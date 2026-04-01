import os

from textual.app import App, ComposeResult
from textual.containers import Container
from textual.widgets import Button, RadioButton, RadioSet, Select, Static

from xls_management.ate import PROJECTS


class MsgBox(App):
    CSS_PATH = 'msgbox.tcss'

    def __init__(self, message:str):
        super().__init__()
        self.message = message

    def compose(self) -> ComposeResult:
        # Create a combobox-like Select widget
        button_ok = Button(label='Ok', id='ok')
        yield Container(
            Container(
                button_ok,
                id='buttons',
            ),
            Container(
                Static(self.message),
                id='msg-container',
            ),
            id='combo-container',
        )

    async def on_button_pressed(self, event:Button.Pressed) -> None:
        '''Handle button pressed event.'''
        #self.ok_clicked = event.button.id == 'ok'
        await self.action_quit()
    
def msgbox(message:str) -> None:
    msgbox = MsgBox(message)
    msgbox.run()