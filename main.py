from kivy.app import App
from kivy.uix.boxlayout import BoxLayout

class MainUI(BoxLayout):
    def calculate(self):
        try:
            a = int(self.ids.a.text)
            b = int(self.ids.b.text)
            c = int(self.ids.c.text)
            d = int(self.ids.d.text)

            total = 0
            steps = ""
            step = 1

            while c > 0:
                value = a * b
                total += value
                steps += f"Step {step}: {a} × {b} = {value}\n"

                a -= 1
                b -= 1
                c -= 1
                step += 1

            total += d

            self.ids.result.text = f"Result: {total}"
            self.ids.steps.text = steps

        except:
            self.ids.result.text = "Error"

class ConeApp(App):
    def build(self):
        return MainUI()

if __name__ == "__main__":
    ConeApp().run()
