from xls_management.shell.ate import MyShell


def test_myshell():
    sh = MyShell()
    sh.cmdloop()
