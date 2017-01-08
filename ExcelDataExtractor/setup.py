from cx_Freeze import setup, Executable
  
setup(
	name = "excel_export", 
	version = "0.1",
	description = "excel_export",
	executables = [Executable("excel_export.py")],
	options = {
		"build_exe" : {
			"create_shared_zip" : False,
			"append_script_to_exe" : True,			
		}
	}
)