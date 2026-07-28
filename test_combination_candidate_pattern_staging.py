from build_combination_candidate_pattern_staging import classify


def row(name: str, mid: str, small: str = "") -> dict[str, str]:
    return {
        "product_name": name,
        "mid_category": mid,
        "small_category": small,
    }


def test_model_suffix_plus_is_not_combination() -> None:
    code, _ = classify(
        row("한샘 아카이브+ 78cm 보조수납장", "책장·수납장")
    )
    assert code == "N01_MODEL_SUFFIX_PLUS"


def test_fullwidth_plus_bed_panel_is_combination() -> None:
    code, _ = classify(
        row("베른 가죽침대 Q ＋ 패널 60cm 2EA", "침대", "침대+매트리스")
    )
    assert code == "Y03_BED_PANEL_GUARD"


def test_bed_and_mattress_is_combination() -> None:
    code, _ = classify(
        row("수납침대 Q + 유로탑 매트리스 Q", "침대", "침대+매트리스")
    )
    assert code == "Y04_BED_SLEEP_SYSTEM"


def test_table_and_chairs_is_combination() -> None:
    code, _ = classify(
        row("1600폭 테이블 + 패브릭 의자 4EA", "식탁·테이블", "식탁")
    )
    assert code == "Y07_TABLE_CHAIR_BENCH"


def test_color_material_join_is_not_combination() -> None:
    code, _ = classify(
        row(
            "체어 그레인 마론 23 + 그레인 네그로 90",
            "의자",
            "인테리어의자",
        )
    )
    assert code == "N04_COLOR_MATERIAL_JOIN"


def test_internal_configuration_stays_review() -> None:
    code, _ = classify(
        row("쿠시노 950폭 수납장(PL박스+책장형)", "유아동가구", "수납장·서랍장")
    )
    assert code == "R01_INTERNAL_CONFIGURATION"
