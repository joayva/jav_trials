from textual import events
from textual.app import App, ComposeResult
from textual.widgets import DataTable, Footer, Header

from xls_management.ate import PROJECTS
from xls_management.com.workbook import Workbook


class TableApp(App):
    """A simple Textual app with a DataTable widget."""

    CSS_PATH = None  # No external CSS for this example

    def compose(self) -> ComposeResult:
        """Create child widgets for the app."""
        yield Header(show_clock=True)
        yield DataTable()
        yield Footer()

    def on_mount(self) -> None:
        """Called when the app is ready."""
        table: DataTable = self.query_one(DataTable)

        # Define columns
        table.add_columns("ID", "Name")

        # Add some rows
        for index, value in enumerate(PROJECTS):
            table.add_row(f"{index}", value)
        # Enable row selection
        table.cursor_type = "row"
        table.focus()

    def on_data_table_row_selected(self, event: DataTable.RowSelected) -> None:
        """Handle row selection event."""
        row_data = event.row_key
        self.log(f"Selected row key: {row_data}")

    async def on_key(self, event: events.Key) -> None:
        """Handle key presses."""
        if event.key == "q":
            await self.action_quit()

def projects():
    TableApp().run()

if __name__ == "__main__":
    TableApp().run()