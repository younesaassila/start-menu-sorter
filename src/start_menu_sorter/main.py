import click
import os
import shutil
import win32com.client
from send2trash import send2trash

start_menu_dirs = [
    os.path.join(os.environ["APPDATA"], "Microsoft", "Windows", "Start Menu"),
    os.path.join(os.environ["PROGRAMDATA"], "Microsoft", "Windows", "Start Menu"),
]
exclude_folders = [
    "Steam",
    "Thrustmaster",
    # Windows
    "Accessibility",
    "Accessories",
    "Administrative Tools",
    "PowerShell",
    "Startup",
    "System Tools",
]


def sort_start_menu(start_menu_dir: str, dry_run: bool):
    # Move all items in root directory to Programs folder
    for item in os.listdir(start_menu_dir):
        path = os.path.join(start_menu_dir, item)
        if os.path.isdir(path):
            if item == "Programs":
                continue
            if not dry_run:
                shutil.move(path, os.path.join(start_menu_dir, "Programs", item))
            print(f"Moved {path} to {os.path.join(start_menu_dir, 'Programs', item)}")
        else:
            if not item.endswith(".lnk"):
                continue
            if not dry_run:
                shutil.move(path, os.path.join(start_menu_dir, "Programs"))
            print(f"Moved {path} to {os.path.join(start_menu_dir, 'Programs')}")

    # Move all items in subfolders to Programs folder
    programs_dir = os.path.join(start_menu_dir, "Programs")
    subfolders = [
        f.path
        for f in os.scandir(programs_dir)
        if f.is_dir()
        and not f.name in exclude_folders
        and not f.name.startswith("Windows")
    ]
    for folder in subfolders:
        shortcuts = [
            os.path.join(root, f)
            for root, dirs, files in os.walk(folder)
            for f in files
            if f.endswith(".lnk")
            and not "uninstall" in f.lower()  # Uninstall shortcuts are a special case
        ]
        if len(shortcuts) <= 1:
            for shortcut in shortcuts:
                # Move shortcut to parent directory
                new_path = os.path.join(programs_dir, os.path.basename(shortcut))
                if not dry_run:
                    shutil.move(shortcut, new_path)
                print(f"Moved {shortcut} to {new_path}")
            # Delete folder
            if not dry_run:
                send2trash(folder)
            print(f"Deleted {folder}")


def clean_shortcuts(start_menu_dir: str, dry_run: bool):
    programs_dir = os.path.join(start_menu_dir, "Programs")
    shortcuts = [
        f.path
        for f in os.scandir(programs_dir)
        if f.is_file() and f.name.endswith(".lnk")
    ]
    for shortcut in shortcuts:
        shell = win32com.client.Dispatch("WScript.Shell")
        shell_shortcut = shell.CreateShortCut(shortcut)
        if not os.path.exists(shell_shortcut.Targetpath):
            if not dry_run:
                send2trash(shortcut)
            print(f"Deleted {shortcut}")


@click.command()
@click.option(
    "--clean",
    "-c",
    is_flag=True,
    help="Clean up shortcuts that point to non-existent files.",
)
@click.option(
    "--dry-run",
    "-d",
    is_flag=True,
    help="Dry run. Don't actually delete anything.",
)
def cli(clean: bool, dry_run: bool):
    try:
        for dir in start_menu_dirs:
            sort_start_menu(dir, dry_run)

            # Clean up invalid shortcuts
            if clean:
                clean_shortcuts(dir, dry_run)

    except PermissionError:
        print("Permission denied. Run as administrator.")


if __name__ == "__main__":
    cli()
