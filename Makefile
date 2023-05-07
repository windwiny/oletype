
oletype-0.4.0-py3-none-any.whl:  oletype/excel.pyi
	pip wheel ./
	.echo pip install oletype-0.4.0-py3-none-any.whl --force-reinstall

oletype/excel.pyi:	gen_win32com.py excel.info.json
	python gen_win32com.py > oletype\excel.pyi

excel.info.json:	downapi.rb
	ruby downapi.rb
