from build_product_component_staging import (
    COMBINATION_CANDIDATE,
    COMBINATION_CONFIRMED,
    COMPONENT_API_CONFIRMED,
    NOT_COMBINATION,
    detect_combination,
    labeled_components_from_text,
    title_components,
)


def test_detect_sofa_stool_combination() -> None:
    status, is_combination, rule = detect_combination(
        "[by CASAMIA] 벨로씨 레보르 천연면피 가죽 소파 4인＋스툴",
        "소파",
    )
    assert status == COMBINATION_CONFIRMED
    assert is_combination == 1
    assert rule == "TITLE_SOFA_PLUS_STOOL"


def test_detect_set_and_model_plus() -> None:
    assert detect_combination("벤치+테이블 세트", "아웃도어가구")[0] == (
        COMBINATION_CONFIRMED
    )
    assert detect_combination("한샘 아카이브+ 78cm 수납장", "책장·수납장")[0] == (
        COMBINATION_CANDIDATE
    )
    assert detect_combination("일반 수납장", "책장·수납장")[0] == NOT_COMBINATION


def test_bed_frame_and_mattress_is_one_3d_asset() -> None:
    status, is_combination, rule = detect_combination(
        "한샘 샘베딩 베이직 침대 SS+노뜨 매트리스",
        "침대",
        "침대+매트리스",
    )
    assert status == NOT_COMBINATION
    assert is_combination == 0
    assert rule == "POLICY_BED_FRAME_MATTRESS_SINGLE_3D_ASSET"


def test_multi_bed_bundle_stays_separate_review_pattern() -> None:
    status, _, _ = detect_combination(
        "저상형 침대 범퍼형 Q+SS 패밀리, 데이베드+노뜨컴포트",
        "침대",
        "침대+매트리스",
    )
    assert status == COMBINATION_CANDIDATE


def test_parse_labeled_component_dimensions() -> None:
    rows = labeled_components_from_text(
        "[4인] W2910 X D1020 X H910mm / [스툴] W740 X D660 X H410mm",
        "소파",
        "벨로씨 레보르 소파 4인＋스툴",
        "test",
    )
    assert [
        (
            row.component_name,
            row.component_type,
            row.w_mm,
            row.d_mm,
            row.h_mm,
            row.resolution_status,
        )
        for row in rows
    ] == [
        ("4인 소파", "SOFA", 2910.0, 1020.0, 910.0, COMPONENT_API_CONFIRMED),
        ("스툴", "STOOL", 740.0, 660.0, 410.0, COMPONENT_API_CONFIRMED),
    ]


def test_parse_parenthesized_axis_labels() -> None:
    rows = labeled_components_from_text(
        "소파 (W)2650 x (D)1020 x (H)760mm / "
        "스툴 (W)800 x (D)600 x (H)450mm",
        "소파",
        "루엔 패브릭 4인 소파+스툴",
        "test",
    )
    assert [
        (row.component_type, row.w_mm, row.d_mm, row.h_mm)
        for row in rows
    ] == [
        ("SOFA", 2650.0, 1020.0, 760.0),
        ("STOOL", 800.0, 600.0, 450.0),
    ]


def test_sofa_title_components_are_two_rows() -> None:
    rows = title_components(
        "벨로씨 레보르 가죽 소파 4인＋스툴",
        "소파",
        "TITLE_SOFA_PLUS_STOOL",
    )
    assert [(row.component_name, row.component_type) for row in rows] == [
        ("4인 소파", "SOFA"),
        ("스툴", "STOOL"),
    ]
