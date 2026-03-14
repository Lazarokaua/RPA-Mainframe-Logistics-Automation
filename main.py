import sys
import os

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), 'src')))

from src.calculo_cubagem import App

def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
