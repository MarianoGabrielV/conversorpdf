from view import App
from controller import Controller

def main():
    app = App()
    controller = Controller(app)
    app.set_controller(controller)
    app.ejecutar()

if __name__ == "__main__":
    main()
