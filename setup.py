from setuptools import setup, find_packages

setup(
    name="start-menu-sorter",
    version="1.0.1",
    description="Sorts the Windows start menu",
    url="https://github.com/younesaassila/start-menu-sorter",
    author="Younes Aassila",
    package_dir={"": "src"},
    packages=find_packages(where="src"),
    python_requires=">=3.8, <4",
    install_requires=["click"],
    entry_points={
        "console_scripts": [
            "start-menu-sorter=start_menu_sorter.main:cli",
        ],
    },
)
