from pathlib import Path

from textual import work
from textual.app import App, ComposeResult
from textual.message import Message
from textual.widgets import Select, Static
from textual_fspicker import FileOpen, Filters


class FilePickerApp(App[None]):
    def __init__(self, path:Path, filters:Filters, title:str="Open"):
        super().__init__()
        self.path = path
        self.file_name:str|None = None
        self.filters = filters
        self.title = title

    @work
    async def on_mount(self) -> None:
        file = FileOpen(
            location=self.path,
            filters=self.filters,
            title=self.title
        )
        result = await self.push_screen_wait(file)
        if result == 'None':
            self.file_name = None
        else:
            self.file_name = result
        await self.action_quit()

def path_from_file_picker(
    location:str=".", 
    filters:Filters=Filters(
        ("Excel", lambda p: p.suffix.lower() == ".xlsx"),
        ("CSV", lambda p: p.suffix.lower() == ".csv"),
        ("All", lambda _: True),
    ),
    title:str="Open"
)-> str:
    file_picker = FilePickerApp(path=location, filters=filters,title=title)
    file_picker.run()
    return file_picker.file_name

if __name__ == "__main__":
    file_name = FilePickerApp(Path(".")).run()
