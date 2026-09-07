"""Sanity tests for package initialization and imports."""

import importlib
import pytest

import scm_allocation


def test_package_version():
    """Verify package version is defined."""
    assert scm_allocation.__version__ == "0.1.0"


@pytest.mark.parametrize(
    "subpackage",
    [
        "scm_allocation.models",
        "scm_allocation.ingestion",
        "scm_allocation.parsers",
        "scm_allocation.normalization",
        "scm_allocation.validation",
        "scm_allocation.output",
        "scm_allocation.config",
        "scm_allocation.utils",
    ],
)
def test_subpackages_importable(subpackage: str):
    """Verify all planned subpackages are properly initialized and importable."""
    module = importlib.import_module(subpackage)
    assert module is not None
