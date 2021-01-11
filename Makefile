make:
	python3 gen.py
install:
	pyinstaller gen.py applications.py date_lib.py docx_lib.py gui_lib.py logic_n_alert.py tool_lib.py --onefile --windowed
	rm -r build
	rm -r gen.spec
