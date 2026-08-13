from pathlib import Path

import pytest

from ppt_quote_linker.cli import main


CASE = Path(__file__).parent / "examples" / "case_01_basic"


def test_cli_creates_output(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    output = tmp_path / "linked.pptx"

    assert main([str(CASE / "input.pptx"), str(output)]) == 0

    assert output.is_file()
    assert "with 24 hyperlinks" in capsys.readouterr().out


def test_cli_rejects_missing_input(tmp_path: Path) -> None:
    with pytest.raises(SystemExit, match="input file does not exist"):
        main([str(tmp_path / "missing.pptx"), str(tmp_path / "output.pptx")])


def test_cli_rejects_overwriting_input() -> None:
    input_path = CASE / "input.pptx"
    with pytest.raises(SystemExit, match="input and output paths must be different"):
        main([str(input_path), str(input_path)])
