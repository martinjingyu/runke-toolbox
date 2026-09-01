from core.module_registry import discover_modules


def test_logistics_module_is_discovered():
    ids = [m.id for m in discover_modules()]
    assert "logistics" in ids
