import argparse
import hashlib
import os
import subprocess
import sys
from pathlib import Path

from rich import print as rprint


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def build_command(
    source: Path,
    build_dir: Path,
    target_exe: Path,
    console_mode: str,
    extra_args: list[str],
) -> list[str]:
    output_filename = os.path.relpath(target_exe, build_dir)
    return [
        sys.executable,
        "-m",
        "nuitka",
        "--mode=onefile",
        f"--windows-console-mode={console_mode}",
        "--enable-plugin=tk-inter",
        f"--output-dir={build_dir}",
        f"--output-filename={output_filename}",
        str(source),
        *extra_args,
    ]


def run_and_capture(command: list[str], cwd: Path, log_path: Path) -> int:
    env = os.environ.copy()
    env.setdefault("PYTHONIOENCODING", "utf-8")

    rprint("[bold cyan]Nuitka command:[/bold cyan]")
    rprint(" ".join(f'"{p}"' if " " in p else p for p in command))
    rprint(f"[dim]Log file: {log_path}[/dim]")

    with log_path.open("w", encoding="utf-8", errors="replace") as log_file:
        proc = subprocess.Popen(
            command,
            cwd=str(cwd),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1,
            env=env,
        )

        assert proc.stdout is not None
        for line in proc.stdout:
            line = line.rstrip("\r\n")
            log_file.write(line + "\n")
            rprint(f"[dim]nuitka>[/dim] {line}")

        return proc.wait()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build clicker.py with Nuitka and verify the exe changed.",
    )
    parser.add_argument("--source", default="clicker.py", help="Python entry file to compile.")
    parser.add_argument("--exe-name", default="TheBestAutoClickerOAT.exe", help="Final exe filename.")
    parser.add_argument("--dist-dir", default="dist", help="Folder for final exe output.")
    parser.add_argument("--build-dir", default="build", help="Folder for Nuitka build artifacts.")
    parser.add_argument(
        "--console-mode",
        default="attach",
        choices=("force", "disable", "attach", "hide"),
        help="Nuitka windows console mode.",
    )
    parser.add_argument(
        "--allow-same-hash",
        action="store_true",
        help="Allow successful build even if output exe hash did not change.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print the Nuitka command without compiling.",
    )
    parser.add_argument(
        "--extra-arg",
        action="append",
        default=[],
        help="Additional Nuitka argument (repeatable).",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    repo_root = Path(__file__).resolve().parent
    source = (repo_root / args.source).resolve()
    build_dir = (repo_root / args.build_dir).resolve()
    dist_dir = (repo_root / args.dist_dir).resolve()
    target_exe = dist_dir / args.exe_name
    log_path = build_dir / "nuitka-build.log"

    if not source.exists():
        rprint(f"[bold red]Error:[/bold red] source file not found: {source}")
        return 2

    build_dir.mkdir(parents=True, exist_ok=True)
    dist_dir.mkdir(parents=True, exist_ok=True)

    old_hash = None
    if target_exe.exists():
        old_hash = sha256_file(target_exe)
        rprint(f"[yellow]Existing exe hash:[/yellow] {old_hash}")
    else:
        rprint("[yellow]No existing exe found; this will be a fresh build.[/yellow]")

    command = build_command(
        source=source,
        build_dir=build_dir,
        target_exe=target_exe,
        console_mode=args.console_mode,
        extra_args=args.extra_arg,
    )

    if args.dry_run:
        rprint("[bold cyan]Dry run mode, compile not executed.[/bold cyan]")
        rprint(" ".join(f'"{p}"' if " " in p else p for p in command))
        return 0

    rc = run_and_capture(command, cwd=repo_root, log_path=log_path)
    if rc != 0:
        rprint(f"[bold red]Build failed.[/bold red] Nuitka exit code: {rc}")
        return rc

    if not target_exe.exists():
        rprint(f"[bold red]Build failed.[/bold red] Output exe not found: {target_exe}")
        return 3

    new_hash = sha256_file(target_exe)
    rprint(f"[green]New exe hash:[/green] {new_hash}")

    if old_hash is not None and new_hash == old_hash and not args.allow_same_hash:
        rprint(
            "[bold red]Build verification failed.[/bold red] "
            "Compilation finished but output exe hash did not change."
        )
        rprint("[dim]Use --allow-same-hash if you want to permit this.[/dim]")
        return 4

    rprint(f"[bold green]Build complete:[/bold green] {target_exe}")
    rprint(f"[bold green]Nuitka output log:[/bold green] {log_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
