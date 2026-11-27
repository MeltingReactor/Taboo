@echo off
python taboo_generator.py taboo_sections.json
echo Done.
timeout /NOBREAK /T 5
exit
