from __future__ import annotations

OPERATOR_MODULE_NAME = "modReplacementOperator"
DAILY_INPUT_MODULE_NAME = "modDailyInputAction"

_MANUAL_MARKER = "    mDecisionAccepted = True\r\n    ConfirmReplacementOperatorChoice = True"
_MANUAL_REPLACEMENT = (
    "    If Not SyncDailyReplacementLedgerFromFacts() Then GoTo ConfirmExit\r\n"
    "    mDecisionAccepted = True\r\n"
    "    ConfirmReplacementOperatorChoice = True"
)

_ROUTE_MARKER = (
    "        If ResolveReplacementWorkflow() Then\r\n"
    "            currentPhase = \"refresh after replacements\""
)
_ROUTE_REPLACEMENT = (
    "        If ResolveReplacementWorkflow() Then\r\n"
    "            If Not SyncDailyReplacementLedgerFromFacts() Then GoTo CleanUp\r\n"
    "            currentPhase = \"refresh after replacements\""
)


def _read_module(vbproject, module_name: str) -> tuple[object, str]:
    component = vbproject.VBComponents(module_name)
    code_module = component.CodeModule
    source = str(code_module.Lines(1, code_module.CountOfLines))
    return code_module, source


def _replace_exact_once(code_module, source: str, marker: str, replacement: str, *, label: str) -> None:
    if replacement in source:
        return
    if source.count(marker) != 1:
        raise RuntimeError(f"{label} marker was not found exactly once")
    transformed = source.replace(marker, replacement, 1)
    code_module.DeleteLines(1, code_module.CountOfLines)
    code_module.AddFromString(transformed)


def install_replacement_ledger_hooks(vbproject) -> None:
    operator_module, operator_source = _read_module(vbproject, OPERATOR_MODULE_NAME)
    _replace_exact_once(
        operator_module,
        operator_source,
        _MANUAL_MARKER,
        _MANUAL_REPLACEMENT,
        label="manual replacement ledger sync",
    )

    daily_module, daily_source = _read_module(vbproject, DAILY_INPUT_MODULE_NAME)
    _replace_exact_once(
        daily_module,
        daily_source,
        _ROUTE_MARKER,
        _ROUTE_REPLACEMENT,
        label="completed replacement ledger sync",
    )
