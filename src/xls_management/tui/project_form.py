import os

from textual.app import App, ComposeResult
from textual.containers import Container
from textual.widgets import Button, RadioButton, RadioSet, Select, Static

from xls_management.ate import PROJECTS


class ProjectChoice(App):
    CSS_PATH = "project_form.tcss"

    def __init__(self):
        super().__init__()
        self.project_name: str = ""
        self.evalue_master_id = False

    def compose(self) -> ComposeResult:
        # Create a combobox-like Select widget
        no_radiobutton: RadioButton = RadioButton("No", value= True)
        yes_radiobutton: RadioButton = RadioButton("Yes")
        radio_set = RadioSet(
            yes_radiobutton,
            no_radiobutton,
            name="Master ID auswerten",
            id="Master_IDs_auswerten"
        )
        button_ok = Button(label="Ok", id="ok")
        yield Container(
            Static("Choose a project:"),
            Select(
                options=[(p, p) for p in PROJECTS],
                prompt="Select project..."
            ),
            Container(
                Static("Evaluate Master Id"),
                Container(
                    radio_set,
                    button_ok,
                    id="choice",
                ),
                id="buttons",
            ),
            id="combo-container"
        )

    async def on_select_changed(self, event: Select.Changed) -> None:
        """Handle selection change event."""
        self.project_name = event.value

    async def on_radio_set_changed(self, event: RadioSet.Changed) -> None:
        """Handle radioset change event."""
        self.evalue_master_id = event.index==0

    async def on_button_pressed(self, event:Button.Pressed) -> None:
        """Handle button pressed event."""
        await self.action_quit()