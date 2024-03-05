"""
Code browser example.

Run with:

    python code_browser.py PATH
"""

import sys
import os
import configHandler
import calendarHandler
import pdfHandler
import docxHandler

from rich.syntax import Syntax
from rich.traceback import Traceback


from textual.app import App, ComposeResult
from textual.containers import Container, VerticalScroll, ScrollableContainer
from textual.reactive import var
from textual.widgets import DirectoryTree, Footer, Header, Static, Button, TextArea, Checkbox

class GetCalendar(Static):
    def compose(self) -> ComposeResult:
        yield Static("Please select a calendar file or URL.")
        yield Static("The file should be in .ics format.")
        yield Static("The calendar file or URL:")
        yield Button("Select file", id="select_file")
        yield Button("Enter URL", id="select_url")
        yield TextArea("", id="calendar_file")
    
class OutputOptions(Static):
    def compose(self) -> ComposeResult:
        yield Static("Please select the output format.")
        yield Checkbox("PDF", id="pdf")
        yield Checkbox("Word (.docx)", id="word")

class GetFileUI(Static):
    def compose(self) -> ComposeResult:
        path = "./" if len(sys.argv) < 2 else sys.argv[1]

        yield Static("Please select a file.")
        yield DirectoryTree(path, id="tree-view")
    def on_mount(self) -> None:
        self.query_one(DirectoryTree).focus()

    def on_directory_tree_file_selected(
        self, event: DirectoryTree.FileSelected
    ) -> None:
        """Called when the user click a file in the directory tree."""
        event.stop()
        code_view = self.query_one("#code", Static)
        try:
            syntax = Syntax.from_path(
                str(event.path),
                line_numbers=True,
                word_wrap=False,
                indent_guides=True,
                theme="github-dark",
            )
        except Exception:
            code_view.update(Traceback(theme="github-dark", width=None))
            self.sub_title = "ERROR"
        else:
            code_view.update(syntax)
            self.query_one("#code-view").scroll_home(animate=False)
            self.sub_title = str(event.path)

    def action_toggle_files(self) -> None:
        """Called in response to key binding."""
        self.show_tree = not self.show_tree


class CalendarApp(App):
    """Textual UI for generating PDF or Word calendars from a YAML file."""
    BINDINGS = [("q", "quit", "Quit the application")]
    CSS_PATH = "gui_style.tcss"
    def compose(self) -> ComposeResult:
        yield Header()
        yield VerticalScroll(GetFileUI())
        # yield VerticalScroll(GetCalendar())
        # yield ScrollableContainer(OutputOptions())
        yield Footer()
    def save_config(self, setting, value):
        '''Change a setting in the config file.'''
        configHandler.change_config(setting, value)
    def load_config(self):
        '''Load the config file.'''
        return configHandler.load_config()
    def get_setting(self, setting):
        '''Get a setting from the config file.'''
        return configHandler.get_setting(setting)
    def quit(self):
        """Quit the application."""
        self.exit()

if __name__ == "__main__":
    CalendarApp().run()