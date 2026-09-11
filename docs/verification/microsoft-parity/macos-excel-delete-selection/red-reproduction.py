import importlib,subprocess,pytest
m=importlib.import_module("skills.WPSComposer.scripts.msoffice.macos_excel_session")
s=subprocess.check_output(["git","show","5622a1615e6fd22676a3aa410a393503dab57f91:skills/WPSComposer/scripts/msoffice/macos_excel_session.py"],text=True)
exec(compile(s,m.__file__,"exec"),m.__dict__)
raise SystemExit(pytest.main(["tests/msoffice/test_macos_excel_session.py","-q","-k","empty_sheet_remov or remove_empty_sheet"]))
