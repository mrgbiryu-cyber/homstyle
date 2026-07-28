from __future__ import annotations

from dimension_context_normalizer import category_profile, extract_candidates


def candidates(name: str, category: str, text: str) -> list[dict]:
    return extract_candidates(text, product_name=name, small_category=category)


def accepted(rows: list[dict]) -> list[dict]:
    return [row for row in rows if row["decision_status"] != "REJECT"]


def test_delivery_clearance_is_rejected() -> None:
    rows = candidates(
        "플라밍고 리클라이너",
        "리클라이너",
        "배송 안내 엘리베이터 내부 규격 W900×D1200×H2300mm",
    )
    assert rows
    assert all(row["candidate_role"] == "DELIVERY_CLEARANCE" for row in rows)
    assert not accepted(rows)


def test_explicit_dimension_section_is_accepted() -> None:
    rows = candidates(
        "퐁드망 TV장",
        "거실장",
        "DIMENSION W1800×D500×H580 mm",
    )
    best = rows[0]
    assert (best["w_mm"], best["d_mm"], best["h_mm"]) == (1800, 500, 580)
    assert best["decision_status"] == "AUTO_ACCEPT"


def test_ldh_and_wlh_furniture_axis_mapping() -> None:
    ldh = candidates(
        "Leiden 3인 소파",
        "소파",
        "SIZE L2050×D900×H750mm",
    )[0]
    assert (ldh["w_mm"], ldh["d_mm"], ldh["h_mm"]) == (2050, 900, 750)
    assert ldh["normalized_axis_mapping"] == "L,D,H->W,D,H"

    wlh = candidates(
        "벤치+테이블 세트",
        "테이블",
        "제품 사이즈 W90×L57×H70cm",
    )[0]
    assert (wlh["w_mm"], wlh["d_mm"], wlh["h_mm"]) == (900, 570, 700)
    assert wlh["normalized_axis_mapping"] == "W,L,H->W,D,H"
    assert wlh["candidate_role"] == "PRODUCT_DIMENSION"


def test_two_dimensional_category_uses_na_depth() -> None:
    rows = candidates("오아시스 아트웍", "아트웍", "제품 규격 60×60cm")
    best = rows[0]
    assert (best["w_mm"], best["d_mm"], best["h_mm"]) == (600, None, 600)
    assert best["normalized_axis_mapping"] == "2D_PAIR->W,H;D=N/A"
    assert best["decision_status"] == "CATEGORY_NORMALIZED"


def test_korean_bed_axes_allow_parenthetical_note() -> None:
    rows = candidates(
        "에이스침대 패밀리",
        "침대",
        "INFO 규격 가로3,220(헤드포함)*세로2,054*헤드높이1,100mm",
    )
    best = rows[0]
    assert (best["w_mm"], best["d_mm"], best["h_mm"]) == (3220, 2054, 1100)


def test_unlabeled_triple_in_size_section() -> None:
    rows = candidates("페블 소파", "소파", "SIZE 2310×990×650")
    best = rows[0]
    assert (best["w_mm"], best["d_mm"], best["h_mm"]) == (2310, 990, 650)
    assert best["decision_status"] in {"AUTO_ACCEPT", "HUMAN_REVIEW"}


def test_title_number_boosts_matching_candidate() -> None:
    rows = candidates(
        "위시본 1600 벤치_Signature",
        "벤치",
        "SIZE W1600×D350×H450",
    )
    assert rows[0]["product_name_match_score"] >= 20


def test_bed_mattress_is_not_classified_as_two_dimensional_mat() -> None:
    assert (
        category_profile("[에이스침대] 패밀리 침대+매트리스", "침대+매트리스")
        == "FURNITURE_3D"
    )


def test_lineup_other_model_is_rejected_near_candidate() -> None:
    text = (
        "한 눈에 보기 제품명 사이즈 Plato Side Table 400 x 400 x 562(mm) "
        "플라토 라인업 Plato Arc Sofa Table 1480 x 520 x 330(mm)"
    )
    rows = candidates("플라토 사이드 테이블", "사이드테이블", text)
    by_width = {row["w_mm"]: row for row in rows if row["w_mm"] is not None}
    assert by_width[400]["candidate_role"] == "PRODUCT_DIMENSION"
    assert by_width[1480]["candidate_role"] == "LINEUP_OTHER_MODEL"
    assert by_width[1480]["decision_status"] == "REJECT"


def test_guard_component_is_rejected_for_bed_product() -> None:
    rows = candidates(
        "[에이스침대] RAQUEL 침대",
        "침대+매트리스",
        "가드 구성 가드 1100폭 W1120/L62/H890mm",
    )
    best = rows[0]
    assert best["candidate_role"] == "COMPONENT_DIMENSION"
    assert best["decision_status"] == "REJECT"


def test_component_material_after_product_size_does_not_relabel_product() -> None:
    rows = candidates(
        "뉴로건 3인 리클라이너",
        "리클라이너",
        "DETAIL 구성 사이즈 W2060 x D1000 x H900mm [프레임] 판재 E0",
    )
    best = rows[0]
    assert best["candidate_role"] == "PRODUCT_DIMENSION"
    assert (best["w_mm"], best["d_mm"], best["h_mm"]) == (2060, 1000, 900)


def test_unit_does_not_capture_m_from_made() -> None:
    rows = candidates(
        "LABRA 소파테이블 140x70cm",
        "거실테이블",
        "SIZE W800*D800*H360(mm) W1400*D700*H310(mm) MADE IN VIETNAM",
    )
    dimensions = {
        (row["w_mm"], row["d_mm"], row["h_mm"])
        for row in rows
        if row["w_mm"] is not None
    }
    assert (1400, 700, 310) in dimensions


def test_targeted_ocr_character_repairs_multi_option() -> None:
    rows = candidates(
        "더함 찻상",
        "거실테이블",
        "Size (M) W490 x D760 x H41 Omm (니 WI 165 x D600 x H360mm",
    )
    dimensions = {
        (row["option_label"], row["w_mm"], row["d_mm"], row["h_mm"])
        for row in rows
        if row["w_mm"] is not None
    }
    assert ("M", 490, 760, 410) in dimensions
    assert ("L", 1165, 600, 360) in dimensions


def test_lwh_oval_and_dh_pedestal_mapping() -> None:
    oval = candidates(
        "Rythme Oval low table",
        "거실테이블",
        "Size L 1100 x W 900 x H 300 (mm)",
    )[0]
    assert (oval["w_mm"], oval["d_mm"], oval["h_mm"]) == (1100, 900, 300)
    assert oval["normalized_axis_mapping"] == "L,W,H->W,D,H"

    round_table = candidates(
        "Good morning Pedestal table",
        "거실테이블",
        "Size D 450 x H 550 (mm)",
    )[0]
    assert (round_table["w_mm"], round_table["d_mm"], round_table["h_mm"]) == (
        450,
        450,
        550,
    )
    assert round_table["shape_type"] == "ROUND"


def test_ace_bed_ocr_axis_repairs() -> None:
    rows = candidates(
        "[에이스침대] BMA1165 패밀리",
        "침대+매트리스",
        "INFO 규격 2 卜로3,220(해드포함)*세로기054*해드 높011,100mm",
    )
    best = rows[0]
    assert (best["w_mm"], best["d_mm"], best["h_mm"]) == (3220, 2054, 1100)


def test_split_thousands_groups_after_korean_axes_are_joined() -> None:
    rows = candidates(
        "Sealy mattress super single",
        "\uce68\ub300+\ub9e4\ud2b8\ub9ac\uc2a4",
        (
            "PRODUCT DETAIL \uc0ac\uc774\uc988 "
            "\ud3ed1,100 x \uae38\uc7742 000 x \ub192\uc774200 "
            "\ud3ed 1 500 x \uae38\uc774 2 000 x \ub192\uc774 200"
        ),
    )
    dimensions = {
        (row["w_mm"], row["d_mm"], row["h_mm"])
        for row in rows
        if row["w_mm"] is not None
    }
    assert (1100, 2000, 200) in dimensions
    assert (1500, 2000, 200) in dimensions
    assert (10, 20, 2000) not in dimensions


def test_labeled_two_dimensional_axes_ignore_spurious_radius_ocr() -> None:
    rows = candidates(
        "limited edition artwork",
        "\uc561\uc790",
        (
            "\uc791\ud488\ud06c\uae30 \uc0ac\uc774\uc988 W83.0cm x H84.0cm "
            "OCR noise 4r1"
        ),
    )
    best = next(
        row
        for row in rows
        if (row["w_mm"], row["h_mm"]) == (830, 840)
    )
    assert best["d_mm"] is None
    assert best["shape_type"] == "AREA_2D"
    assert best["normalized_axis_mapping"] == "W,H->W,H;D=N/A"


def test_flattened_mattress_table_does_not_carry_ss_to_next_width() -> None:
    rows = candidates(
        "Sealy mattress super single (SS)",
        "\uce68\ub300+\ub9e4\ud2b8\ub9ac\uc2a4",
        (
            "PRODUCT DETAIL \uc0ac\uc774\uc988 "
            "\uc288\ud37c\uc2f1\uae00 SS \ud3ed1,100 x \uae38\uc7742 000 x \ub192\uc774200 "
            "\ud3ed1,500 x \uae38\uc7742 000 x \ub192\uc774200"
        ),
    )
    by_width = {
        row["w_mm"]: row
        for row in rows
        if row["w_mm"] is not None and row["d_mm"] is not None
    }
    assert by_width[1100]["option_label"] == "SS"
    assert by_width[1500]["option_label"] == ""
    assert (
        by_width[1100]["product_name_match_score"]
        > by_width[1500]["product_name_match_score"]
    )


def test_dimension_range_is_not_automatically_collapsed_to_first_value() -> None:
    rows = candidates(
        "Cliff TV Stand",
        "TV stand",
        "SIZE L:1250-2350 / D:380 / H:400 mm",
    )
    best = next(row for row in rows if row["w_mm"] == 1250)
    assert best["decision_status"] == "HUMAN_REVIEW"
    assert best["rejection_reason"] == "DIMENSION_RANGE_REVIEW"


def test_middle_axis_range_keeps_other_axes_and_requires_review() -> None:
    rows = candidates(
        "로이모노 1200폭 5단 책상세트",
        "일자형책상",
        (
            "MODEL SIZE 로이모노 5단 다리형 책상세트 1200폭 "
            "가로 1200 x 세로 718~891 x 높이 1705 mm"
        ),
    )
    best = next(
        row
        for row in rows
        if row["w_mm"] == 1200
        and row["d_mm"] == 718
        and row["h_mm"] == 1705
    )
    assert "718~891" in best["raw_notation"]
    assert best["decision_status"] == "HUMAN_REVIEW"
    assert best["rejection_reason"] == "DIMENSION_RANGE_REVIEW"


def test_zero_area_axis_is_reocr_required() -> None:
    rows = candidates(
        "잭 러그",
        "러그",
        "INTRODUCTION SIZE 00 x 150cm",
    )
    best = next(row for row in rows if row["w_mm"] == 0)
    assert best["decision_status"] == "REOCR_REQUIRED"
    assert best["rejection_reason"] == "OCR 숫자 결합·단위 이상"


def test_local_component_label_excludes_accessory_dimensions() -> None:
    rows = candidates(
        "로이모노 1200폭 5단 책상세트",
        "일자형책상",
        (
            "MODEL SIZE 로이모노 5단 다리형 책상세트 "
            "가로 1200 세로 718 높이 1705 "
            "상부도어 : 가로 1194 세로 18.5 높이 349 "
            "다용도선반 / 다용도꽂이 / 원형 자석 : "
            "가로 109 세로 229 높이 30"
        ),
    )
    accessory = next(
        row
        for row in rows
        if row["w_mm"] == 1194
        and row["d_mm"] == 18.5
        and row["h_mm"] == 349
    )
    assert accessory["candidate_role"] == "COMPONENT_DIMENSION"
    assert accessory["decision_status"] == "REJECT"


def test_model_info_component_word_does_not_reject_matching_twin_frame() -> None:
    rows = candidates(
        "ACE BMA1169 / TW",
        "\uce68\ub300+\ub9e4\ud2b8\ub9ac\uc2a4",
        (
            "INFO \ubaa8\ub378\uba85 \uad6c\uc131 BMA1169 \uce68\ub300 \ud504\ub808\uc784 + "
            "\ub9e4\ud2b8\ub9ac\uc2a4 \uc0ac\uc774\uc988 \uaddc\uaca9 "
            "\ud2b8\uc708 : \uac00\ub85c2,788 x \uc138\ub85c2,066 x "
            "\ud5e4\ub4dc \ub192\uc7741,208mm"
        ),
    )
    best = next(row for row in rows if row["w_mm"] == 2788)
    assert best["candidate_role"] == "PRODUCT_DIMENSION"
    assert best["decision_status"] == "AUTO_ACCEPT"
    assert best["option_label"] == "TW"


def test_ace_bed_info_width_leading_two_read_as_hangul_gi() -> None:
    rows = candidates(
        "ACE BMA1150 침대",
        "\uce68\ub300+\ub9e4\ud2b8\ub9ac\uc2a4",
        (
            "INFO 모델명 BMA1150 규격 "
            "가로기600 * 세로2140 * 헤드높이800mm"
        ),
    )
    best = next(row for row in rows if row["w_mm"] == 2600)
    assert (best["w_mm"], best["d_mm"], best["h_mm"]) == (2600, 2140, 800)


if __name__ == "__main__":
    tests = [
        value
        for key, value in sorted(globals().items())
        if key.startswith("test_") and callable(value)
    ]
    for test in tests:
        test()
        print(f"PASS {test.__name__}")
