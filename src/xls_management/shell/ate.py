import argparse
import cmd
from importlib.metadata import version
from pathlib import Path
import shlex
from typing import Any

from xls_management.ate import PROJECTS
from xls_management.ate.project import project_combo_box
from xls_management.ate.tracking import ATEStatus
from xls_management.tui.file_picker import path_from_file_picker
from xls_management.tui.yes_no_form import yes_no_msgbox
from xls_management.utils.color import Color, ansi_color


class MyShell(cmd.Cmd):
    intro = (
        'Welcome to ate shell. Type help or ? to list commands.\n'
        f'Available projects: {", ".join(PROJECTS)}\n'
    )
    prompt = "(ate) "

    def __init__(self):
        super().__init__()
        self.ate_status = ATEStatus()

    def parse_config_args(self, args) -> Any:
        # Prepare argument parsers for commands
        config_parser = argparse.ArgumentParser(prog="config", add_help=True)
        config_parser.add_argument("file", default="config.yml")
        config_parser.add_argument(
            '-c', 
            '--create',
            action='store_true',
            default= False,
            required= False,
            help='config <file> --create',
        )
        return config_parser.parse_args(args)

    def parse_status_args(self, args) -> Any:
        # Prepare argument parsers for commands
        status_parser = argparse.ArgumentParser(prog="config", add_help=True)
        status_parser.add_argument("file", default="config.yml")
        return status_parser.parse_args(args)

    # Example command
    def do_project(self, arg):
        """project <name>"""
        if arg:
            print(arg)

    def do_status(self, arg):
        """status"""
        
        if arg:
            args = self.parse_config_args(arg.split())
            print(args)
            path = Path(args.file)
            if path.exists():
                self.ate_status.config.config_from(path)
                assert self.ate_status.config is not None
                self.ate_status.perform_status()
            else:
                print(f'{args.file} does not exist!')
        else:
            assert self.ate_status.config is not None
            self.ate_status.perform_status()

    def do_choose(self, arg):
        """choose a project from a list using a combobox"""
        str_project = arg
        evalue_master_id: bool
        if not str_project:
            str_project, evalue_master_id = project_combo_box()
        print(f'project {ansi_color(str_project, Color.GREEN)} was chosen')
        msg = f'{ansi_color("evalue",Color.GREEN)} master id'
        if not evalue_master_id:
            msg = f'{ansi_color("do not evalue",Color.RED)} master id'
        print(msg)

    def do_config(self, arg):
        """choose a config file"""
        if arg:
            args = self.parse_config_args(arg.split())
            print(args)
            path = Path(args.file)
            if args.create:
                print(f"adding config file {args.file} ...")
                self.ate_status.config_to_file(path)
            elif  path.exists():
                self.ate_status.config.config_from(path)
                print(ansi_color(f'Updated config from {arg}', Color.GREEN))
            else:
                print(ansi_color(f'{arg} file does not exist', Color.RED))
        else:
            print(ansi_color('A file path is required', Color.RED))

    def do_ask(self, arg):
        """choose a project from a list using a combobox"""
        yes_answer = yes_no_msgbox('Trial msgbox\nAre you ready to answer?')
        if yes_answer:
            print(f'User answer is {ansi_color("YES", Color.GREEN)}')
        else:
            print(f'User answer is {ansi_color("NO",Color.RED)}')
    
    def do_file(self, location:str="."):
        """uses a file picker to choose a file"""
        file_path = path_from_file_picker(location=location, title='Choose ATE input file')
        print(f'file {ansi_color(str(file_path),Color.GREEN)} was chosen')

    def do_reset(self, arg):
        """resets the config file to default values"""
        ate_status:ATEStatus|None = ATEStatus()
        if yes_no_msgbox('This will remove. Do you want to continue?'):
            ate_status.erase()
            print('Config file reset to default values.')
        else:
            print('Config file reset cancelled.')
    # Tab completion for 'project'
    def complete_project(self, text, line, begidx, endidx):
        """
        text: the current word being completed
        line: the full input line
        begidx, endidx: start and end indexes of the word
        """
        if not text:
            return PROJECTS  # Suggest all if nothing typed yet
        return [name for name in PROJECTS if name.lower().startswith(text.lower())]

    # Exit command
    def do_exit(self, arg):
        """Exit the shell"""
        print('Goodbye!')
        return True

def launch_shell():
    parser = argparse.ArgumentParser(description='customized shell')
    parser.add_argument('-v', '--version', action='store_true')
    args = parser.parse_args()
    if args.version:
        print(version('xls_management'))
    else:
        MyShell().cmdloop()

if __name__ == "__main__":
    launch_shell()
