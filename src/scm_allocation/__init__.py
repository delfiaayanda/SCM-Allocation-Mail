"""SCM Allocation Mail package.

Automated extraction, normalization, and validation of SCM allocation requests.
"""

__version__ = "0.1.0"

from scm_allocation.extraction import extract_email, is_out_of_scope_email

__all__ = ["__version__", "extract_email", "is_out_of_scope_email"]
