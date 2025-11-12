"""Command-line interface for slidecat."""

import sys
from pathlib import Path
from typing import Optional

import click

from .core import split_presentation, merge_presentations, extract_slides, verify_presentation


@click.group()
@click.version_option(version="0.1.0")
def main():
    """slidecat - Split and merge PowerPoint files."""
    pass


@main.command()
@click.argument('input_file', type=click.Path(exists=True, path_type=Path))
@click.option(
    '-o', '--output-dir',
    type=click.Path(path_type=Path),
    default=Path('./slides'),
    help='Output directory for split files (default: ./slides)'
)
@click.option(
    '-c', '--chunk-size',
    type=int,
    default=None,
    help='Number of slides per file (default: 1 slide per file)'
)
def split(input_file: Path, output_dir: Path, chunk_size: Optional[int]):
    """
    Split a PowerPoint file into individual slides or chunks of slides.

    INPUT_FILE: Path to the PPTX file to split
    """
    try:
        if chunk_size is not None:
            click.echo(f"Splitting {input_file} into chunks of {chunk_size} slides...")
        else:
            click.echo(f"Splitting {input_file}...")

        created_files = split_presentation(input_file, output_dir, chunk_size)

        click.echo(f"✓ Successfully split into {len(created_files)} files")
        click.echo(f"  Output directory: {output_dir.absolute()}")

        if len(created_files) <= 10:
            for file in created_files:
                click.echo(f"    - {file.name}")
        else:
            click.echo(f"    - {created_files[0].name}")
            click.echo(f"    - ...")
            click.echo(f"    - {created_files[-1].name}")

    except Exception as e:
        click.echo(f"Error: {e}", err=True)
        sys.exit(1)


@main.command()
@click.argument('input_files', nargs=-1, required=True, type=click.Path(exists=True, path_type=Path))
@click.option(
    '-o', '--output',
    type=click.Path(path_type=Path),
    required=True,
    help='Output file path'
)
def merge(input_files: tuple, output: Path):
    """
    Merge multiple PowerPoint files into one.

    INPUT_FILES: Paths to PPTX files to merge (space-separated)
    """
    try:
        input_list = list(input_files)
        click.echo(f"Merging {len(input_list)} files...")

        for file in input_list:
            click.echo(f"  - {Path(file).name}")

        output_file = merge_presentations(input_list, output)

        click.echo(f"✓ Successfully merged into {output_file.absolute()}")

    except Exception as e:
        click.echo(f"Error: {e}", err=True)
        sys.exit(1)


@main.command()
@click.argument('input_file', type=click.Path(exists=True, path_type=Path))
@click.option(
    '-o', '--output',
    type=click.Path(path_type=Path),
    required=True,
    help='Output file path'
)
@click.option(
    '-r', '--range',
    'slide_range',
    required=True,
    help='Slide range to extract (e.g., "1-5" or "3-")'
)
def extract(input_file: Path, output: Path, slide_range: str):
    """
    Extract a range of slides from a PowerPoint file.

    INPUT_FILE: Path to the PPTX file
    """
    try:
        # Parse the range
        if '-' not in slide_range:
            click.echo("Error: Range must be in format 'start-end' or 'start-'", err=True)
            sys.exit(1)

        parts = slide_range.split('-', 1)
        start = int(parts[0])
        end = int(parts[1]) if parts[1] else None

        click.echo(f"Extracting slides {start} to {end or 'end'} from {input_file}...")

        output_file = extract_slides(input_file, output, start, end)

        click.echo(f"✓ Successfully extracted to {output_file.absolute()}")

    except ValueError as e:
        click.echo(f"Error: Invalid range format - {e}", err=True)
        sys.exit(1)
    except Exception as e:
        click.echo(f"Error: {e}", err=True)
        sys.exit(1)


@main.command()
@click.argument('input_file', type=click.Path(exists=True, path_type=Path))
@click.option(
    '-v', '--verbose',
    is_flag=True,
    help='Show detailed information for each slide'
)
def verify(input_file: Path, verbose: bool):
    """
    Verify a PowerPoint file and check for errors.

    INPUT_FILE: Path to the PPTX file to verify
    """
    try:
        click.echo(f"Verifying {input_file}...")

        result = verify_presentation(input_file)

        if result["valid"]:
            click.echo(f"✓ File is valid")
            click.echo(f"  Total slides: {result['slides']}")
            click.echo(f"  Slide layouts: {result['layouts']}")
            click.echo(f"  Slide masters: {result['masters']}")

            if verbose:
                click.echo("\n  Slide details:")
                for slide in result["slide_details"]:
                    status = "✓" if not slide["error"] else "✗"
                    click.echo(f"    {status} Slide {slide['number']}: {slide['shapes']} shapes", nl=False)
                    if slide["has_title"]:
                        click.echo(" (has title)", nl=False)
                    if slide["error"]:
                        click.echo(f" - ERROR: {slide['error']}", nl=False)
                    click.echo()

            # Check for any slide errors
            error_slides = [s for s in result["slide_details"] if s["error"]]
            if error_slides:
                click.echo(f"\n  Warning: {len(error_slides)} slide(s) have errors")
                for slide in error_slides:
                    click.echo(f"    - Slide {slide['number']}: {slide['error']}")
        else:
            click.echo(f"✗ File is invalid", err=True)
            click.echo(f"  Error: {result['error']}", err=True)
            sys.exit(1)

    except Exception as e:
        click.echo(f"Error: {e}", err=True)
        sys.exit(1)


if __name__ == '__main__':
    main()
