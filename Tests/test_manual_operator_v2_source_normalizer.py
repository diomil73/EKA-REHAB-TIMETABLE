from rehab_excel.manual_operator_v2_source_normalizer import normalize_vba_source
from rehab_excel.manual_replacement_operator_v2_patch import build_v2_module_code
from rehab_excel.replacement_operator_vba import MODULE_CODE


def test_normalizer_converts_excel_com_crlf_before_v2_patch():
    crlf_source = MODULE_CODE.replace("\n", "\r\n")
    normalized = normalize_vba_source(crlf_source)

    assert "\r" not in normalized
    patched = build_v2_module_code(normalized)
    assert "MANUAL_OPERATOR_V2" in patched
    assert "RequestNextReplacementPatient" in patched
    assert "RequestAutomaticRemainingReplacements" in patched
