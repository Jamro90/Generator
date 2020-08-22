# compilating command
make:
	python3 generator.py
# clearing command
remove:
	rm -r build dist generator.spec
# making end product
install:
	pyinstaller generator.py --onefile --icon=icons/icon.png -w

