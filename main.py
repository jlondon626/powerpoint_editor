"""Example entrypoint for the xmlppt package.

This file provides a small CLI to exercise the PowerPointEditor when
an input presentation is available. For programmatic usage prefer
importing `xmlppt.PowerPointEditor` from your code.
"""

from xmlppt import PowerPointEditor
import argparse
import os
import sys


def run_example(input_pptx: str, output_pptx: str | None = None) -> str:
    editor = PowerPointEditor(input_pptx=input_pptx)

    editor.list_sections()
    editor.list_all_textboxes()
    editor.list_graphic_frames()

    try:
        slide = editor.duplicate_template_slide(
            template_name="RESERVE_WATERFALL",
            before_section_name="template_slides",
        )

        slide.edit_textbox_html(
            textbox_name="AOM Text",
            html=(
                "This is a generated slide for <b>Class A</b>.<br>"
                "The text has been edited after duplication."
            ),
        )

        slide.edit_embedded_workbook_for_chart(
            chart_name="Waterfall chart",
            categories=[
                "2024 Q4",
                "AvE",
                "Change in\nPremium",
                "Specifics",
                "CATs",
                "2025 Q4",
            ],
            values=[1000, -120, 80, -14, -25, 921],
            subtotal_indices=[0, 5],
        )

    except Exception as exc:
        print(f"Example actions skipped due to error: {exc}")

    output = editor.save(output_pptx) if output_pptx else editor.save()
    return output


def main(argv: list[str] | None = None) -> int:
    argv = argv if argv is not None else sys.argv[1:]
    parser = argparse.ArgumentParser(description="xmlppt example CLI")
    parser.add_argument("--input", "-i", help="Input .pptx file to operate on")
    parser.add_argument("--output", "-o", help="Output .pptx file (optional)")
    parser.add_argument("--run-example", action="store_true", help="Run the bundled example operations")

    args = parser.parse_args(argv)

    if not args.input:
        parser.print_help()
        print("\nExample: python main.py --input example.pptx --run-example")
        return 1

    if not os.path.exists(args.input):
        print(f"Input file not found: {args.input}")
        return 2

    if args.run_example:
        out = run_example(args.input, args.output)
        print(f"Created: {out}")
    else:
        print("No action requested; use --run-example to run the sample flow.")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
