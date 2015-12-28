import tkinter as tk


class ShowMessage:
    def __init__(self, message, interval=4000):
        self.root = tk.Tk()
        text = tk.Text(self.root, width=40, height=5)
        text.insert(1.0, message)
        text.pack()
        self.root.after_idle(self.close)
        self.message = message
        self.interval = interval

    def close(self):
        self.root.after(self.interval, self.root.destroy)

    def start(self):
        self.root.mainloop()


if __name__ == "__main__":
    show = ShowMessage("hello, world")
    show.start()