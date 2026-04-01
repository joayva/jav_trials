import os

from textual.app import App, ComposeResult
from textual.containers import Container
from textual.widgets import Button, RadioButton, RadioSet, Select, Static

from xls_management.ate import PROJECTS


class YesNoChoice(App):
    CSS_PATH = "yes_no_form.tcss"

    def __init__(self, message:str):
        super().__init__()
        self.message = message
        self.yes_clicked = False

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
        button_no = Button(label="No", id="no")
        button_yes = Button(label="Yes", id="yes")
        yield Container(
            Static(self.message),
            Container(
                button_no,
                button_yes,
                id="buttons",
            ),
            id="combo-container"
        )

    async def on_button_pressed(self, event:Button.Pressed) -> None:
        """Handle button pressed event."""
        self.yes_clicked = event.button.id == "yes"
        await self.action_quit()
    
def yes_no_msgbox(message:str) -> bool:
    yes_no = YesNoChoice(message)
    yes_no.run()
    return yes_no.yes_clicked