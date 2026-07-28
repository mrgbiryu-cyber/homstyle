from low_dimension_quality_policy import assess_low_dimension


def test_sofa_depth_12_requires_review() -> None:
    result = assess_low_dimension(
        "알로소 4인 소파",
        "소파",
        "일반소파",
        900,
        12,
        2300,
    )
    assert result.requires_review
    assert result.code == "LOW_DIMENSION_REVIEW_REQUIRED"
    assert result.low_axes == ("D",)


def test_rug_height_10_is_thin_product_exception() -> None:
    result = assess_low_dimension(
        "원형 코튼 러그",
        "러그·매트",
        "러그",
        1500,
        1500,
        10,
    )
    assert not result.requires_review
    assert result.code == "THIN_PRODUCT_EXCEPTION_REQUIRES_EVIDENCE"


def test_picture_frame_depth_16_is_thin_product_exception() -> None:
    result = assess_low_dimension(
        "LEGACY Picture frame",
        "갤러리·월데코",
        "액자",
        250,
        16,
        300,
    )
    assert not result.requires_review


def test_storage_box_height_19_requires_review() -> None:
    result = assess_low_dimension(
        "노블 우드 휴지 케이스",
        "책장·수납장",
        "수납장",
        210,
        261,
        19,
    )
    assert result.requires_review


def test_hanger_depth_4_is_thin_product_exception() -> None:
    result = assess_low_dimension(
        "알루미늄 상의용 옷걸이",
        "옷장·행거",
        "행거",
        420,
        4,
        215,
    )
    assert not result.requires_review


def test_multiple_low_axes_are_not_auto_exception() -> None:
    result = assess_low_dimension(
        "얇은 액자",
        "갤러리·월데코",
        "액자",
        20,
        10,
        300,
    )
    assert result.requires_review

