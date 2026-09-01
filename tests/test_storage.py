from core.storage.local import LocalStorage


def test_local_storage_round_trip(tmp_path):
    storage = LocalStorage(root=str(tmp_path))

    storage.write_bytes("sub/hello.txt", b"hi")

    assert storage.exists("sub/hello.txt")
    assert storage.read_bytes("sub/hello.txt") == b"hi"
    assert storage.list_files("sub") == ["hello.txt"]


def test_local_storage_missing_file(tmp_path):
    storage = LocalStorage(root=str(tmp_path))

    assert not storage.exists("nope.txt")
    assert storage.list_files("nope_dir") == []
