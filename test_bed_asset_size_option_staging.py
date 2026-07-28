from build_bed_asset_size_option_staging import bed_size_codes


def test_korean_and_code_sizes() -> None:
    assert bed_size_codes("슈퍼싱글(SS)") == ["SS"]
    assert bed_size_codes("퀸(Q)") == ["Q"]
    assert bed_size_codes("칼킹(CK)") == ["CK"]


def test_multiple_sizes() -> None:
    assert bed_size_codes("침대 사이즈 택1 Q/K") == ["Q", "K"]


def test_non_size_option() -> None:
    assert bed_size_codes("화이트 루바형") == []
