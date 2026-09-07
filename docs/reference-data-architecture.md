# Reference Data & Routing Architecture

## 1. Executive Summary & Separation of Concerns

The SCM Allocation Mail system strictly decouples **email parsing** from **reference data validation**:

```text
EMAIL INGESTION / PARSING LAYER
(Plain Text, HTML Tables, Excel Attachments)
                │
                ▼
  Unvalidated Candidate Allocation Records
  [WH Code, Plant Code, Material, Qty, ...]
                │
                ▼
┌─────────────────────────────────────────────────────────────┐
│               MASTER DATA & ROUTING VALIDATION              │
│                                                             │
│       ReferenceDataProvider                                 │
│                 │                                           │
│                 ▼                                           │
│          RoutingProvider ◄──────── RoutingSnapshot (Cache)   │
│                 │                                           │
│                 ├── LocalSnapshotRoutingProvider (Offline)  │
│                 └── ApexRoutingProvider (API Sync)          │
└─────────────────────────────────────────────────────────────┘
                │
                ▼
    Validated Allocation Requests + Discrepancy Flags
```

### Why Reference Data Must Be Separated:
1. **Dynamic Volatility vs. Parser Invariance**: Email table structures (rows, columns, formatting) change independently of company master data and distribution routings.
2. **Offline Reproducibility & Testing**: Unit tests for email extraction must run deterministically in local environments without network connections or external service dependencies.
3. **Decoupled Failure Domains**: If an external reference API is temporarily unreachable or slow, incoming emails can still be ingested and queued rather than dropping transactions.
4. **Auditability & Traceability**: Allocation validation decisions must record the exact snapshot version or timestamp (`fetchedAt`) against which the decision was validated.

---

## 2. APEX RICO Routing Sync Discovery

Reconnaissance of the APEX RICO portal's "Daily Sync" button via HTTP Archive (HAR) analysis revealed:

- **Endpoint**: `POST https://api.apexrico.space/api/routing/sync`
- **Method**: `POST` (executed with an empty payload `{}`)
- **Response Format**: JSON containing:
  - `rows`: Array of 2,941 routing records across 31 corporate SAP fields.
  - `log`: Aggregation summary of source ERP partitions:
    - *NASA SAP*: 247 rows
    - *EAR SAP*: 2,220 rows
    - *Ibox SAP*: 323 rows
    - *MII SAP*: 151 rows
    - *Merged Total*: 2,941 rows
  - `fetchedAt`: ISO 8601 timestamp representing the exact generation moment of the snapshot.
  - `hash`: SHA/hash signature verifying dataset integrity.

### Direct API vs. Browser Automation:
Browser automation (Selenium, Playwright, Puppeteer) is **strictly unnecessary** and undesirable for this integration. The endpoint can be queried via lightweight direct HTTP calls or loaded from cached snapshots, avoiding headless browser overhead, driver fragility, and browser binary dependencies.

---

## 3. Reference Data Provider Abstraction

To ensure the core allocation engine never couples to a specific network implementation or vendor URL, we establish an abstract provider hierarchy:

```text
                 ReferenceDataProvider [ABC]
                           │
                           ▼
                  RoutingProvider [ABC]
                           │
             ┌─────────────┴─────────────┐
             ▼                           ▼
   LocalSnapshotRoutingProvider   ApexRoutingProvider
   (Offline / JSON cache)         (Live HTTP Sync)
```

### Core Abstractions:

1. **`ReferenceDataProvider` (Generic ABC)**:
   Base contract for any corporate master data provider (warehouses, plants, materials, routings).

2. **`RoutingProvider` (Domain ABC)**:
   Contract for routing data access:
   - `get_routing_snapshot(force_refresh: bool = False) -> RoutingSnapshot`
   - `resolve_wh_sender(plant_code: str, context: AllocationContext) -> Optional[str]`

3. **`LocalSnapshotRoutingProvider`**:
   Reads a frozen `RoutingSnapshot` from a local JSON file on disk. Used for:
   - Automated offline integration tests.
   - Air-gapped / staging environments.
   - Fallback when network sync fails.

4. **`ApexRoutingProvider`**:
   Connects to the APEX RICO API endpoint. Handles network timeouts, HTTP headers, response decompression, error propagation, and writes through to the local snapshot cache.

---

## 4. Routing Snapshot & Caching Architecture

Querying 2,941 rows across the internet for every individual email is inefficient and introduces latency and external rate-limit risk. Instead, the system operates on **Routing Snapshots**:

```text
  APEX RICO API (Daily Sync)
              │
              ▼ (Once per batch / day)
       RoutingSnapshot
       ├── fetched_at: "2026-09-07T..."
       ├── source: "apex_api"
       ├── hash_code: "..."
       └── rows: [ 2,941 records ]
              │
              ├── Indexed in-memory: dict[plant_code -> list[RoutingRow]]
              ▼
   Process 100+ Allocation Emails Locally (Sub-millisecond lookup)
```

### Snapshot Structure:
- `fetched_at`: Timestamp recording when the snapshot was pulled.
- `source`: Origin identifier (e.g. `apex_api` or `local_fixture`).
- `hash_code`: Data version hash for cache validation.
- `total_rows`: Count of validated rows.
- `rows`: Strongly-typed `RoutingRow` entities.

---

## 5. Multi-Context Allocation Routing

Corporate routing is **not a simple 1-to-1 mapping** from `Plant Code` to `WH Code`. The 31 columns reveal that supply chain fulfillment routes differ by **Product Category / Allocation Context**:

| Context | Forward WH Sender Column | Forward Plant Column | Forward Schedule Column | Typical Scope |
| :--- | :--- | :--- | :--- | :--- |
| **Device** | `WH Sender Code (Alokasi Device)` | `Plant Code (Alokasi Device)` | `JADWAL Device` | Handphones, Tablets, MacBooks |
| **Accessories** | `WH Sender Code (Alokasi Accs)` | `Plant Code (Alokasi Accs)` | `JADWAL Accs` | Chargers, Cases, Belkin, Logitech |
| **NPI** | `WH Sender Code (Alokasi NPI)` | `Plant Code (Alokasi NPI)` | N/A | New Product Introduction launches |
| **TV** | `WH Sender Code (Alokasi TV)` | `Plant Code (Alokasi TV)` | N/A | Television / Smart Screen displays |

### Real-World Email Validation Relevance:
- In **Email 002** (`Alokasi item Logitech dan Belkin to XE05`):
  - Category: Accessories (`Belkin`, `Logitech`).
  - Validation must check `WH Sender Code (Alokasi Accs) == "BIC1"`.
- In **Emails 003, 004, 005** (`iPhone 17`, `MacBook Neo`):
  - Category: Device.
  - Validation must check `WH Sender Code (Alokasi Device) == "BIC1"`.

> [!IMPORTANT]
> **Architectural Decision Rule**:
> The validation layer must accept an `allocation_context` parameter (or infer it from the material category) when cross-referencing valid supply warehouses. A warehouse valid for Accessories may not be the designated routing warehouse for Devices.

---

## 6. Authentication & Security Considerations

1. **Zero Credential Hardcoding**:
   - The HAR capture currently showed no Authorization header on that specific endpoint; however, **we do not assume the API is permanently unauthenticated**.
   - The architecture supports configurable headers via environment variables (`APEX_API_KEY`, `APEX_BEARER_TOKEN`).
2. **No Secret or Payload Commits**:
   - Neither the `.har` capture file nor production routing payloads are committed to version control.
   - `*.har` is explicitly excluded in `.gitignore`.
3. **Defense-in-Depth**:
   - Network errors from the external API must raise explicit typed exceptions (`RoutingSyncError`, `RoutingConnectionError`) and never crash the email ingestion service.

---

## 7. Status Matrix: Confirmed vs. Unknown

### Confirmed:
- [x] API endpoint URL and HTTP method: `POST https://api.apexrico.space/api/routing/sync`.
- [x] Empty request body payload `{}`.
- [x] Response schema: `rows`, `log`, `fetchedAt`, `hash`.
- [x] Total merged row count: 2,941 rows across NASA, EAR, Ibox, and MII SAP systems.
- [x] Distinct routing contexts: Device, Accessories, NPI, and TV.
- [x] No browser automation required.

### Unknown / Pending Verification:
- [ ] Rate limits or throttling on `api.apexrico.space`.
- [ ] Whether IP whitelisting or corporate VPN is required in production.
- [ ] Authentication requirements for production service accounts vs. browser sessions.
- [ ] Handling of stores with multiple valid warehouse senders or custom temporary route overrides.
