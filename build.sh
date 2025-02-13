#!/usr/bin/env bash
PyInstaller --windowed --name="PDF Tools" --icon=pdf-tools.icns --add-data="pdf-tools.icns:." --add-data="icons/book-question.png:icons" --add-data="icons/disk.png:icons" --add-data="icons/folder-horizontal-open.png:icons" --add-data="icons/information-white.png:icons" --add-data="icons/arrow-circle.png:icons" pdf-tools.py
