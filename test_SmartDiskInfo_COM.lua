win32exts.load_sym("*", "*")

SmartDevice = win32exts.create_object("{4cc61604-07eb-4e25-b336-fc1b93e3fb1a}")
--"{1a3aa9dd-5a95-42dd-8a22-0c95af915d3f}") SmartResult

sr = SmartDevice.invoke("get_drive_info", "C:")
sr.add_ref()
dispId = SmartDevice.find_sym("get_drive_info")

dispId = sr.find_sym("root_item")
si = sr.invoke("root_item")
si.add_ref()
dispId = sr.find_sym("root_item")
sn = si.invoke("item_by_key", "serial_number")
sn.add_ref()
strtext = sn.invoke("to_string")

win32exts.MessageBoxW(nil, {strtext}, {"serial_number of C:"}, 0)
