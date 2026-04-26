# -*- coding: utf-8 -*
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.spinner import Spinner

class CustomSpinner(Spinner):
    pass

class MainRoot(BoxLayout):

    def on_spinner_change(self, text):
        print('The spinner', self, 'have text', text)

class MainApp(App):
    def build(self):
        self.title = 'Spinner Test'

if __name__ == "__main__":
    MainApp().run()
