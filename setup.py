from cx_Freeze import setup, Executable

setup(
    name="StickerApp",
    version="1.0",
    description="Generador de etiquetas",
    executables=[Executable("etiquetas.py", base="Win32GUI", icon="sticker.ico")],
    options={
        "build_exe": {
            "packages": ["reportlab", "tkinter", "os", "pandas"],
            "include_files": [],  # puedes incluir fuentes o íconos si es necesario
        }
    }
)
