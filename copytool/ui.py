# -*- coding: utf-8 -*-
import tkinter as tk
import logging

logger = logging.getLogger(__name__)

class ScreenSelector:
    """Class to create a transparent overlay for screen area selection."""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.overrideredirect(True)
        self.root.attributes('-alpha', 0.3)
        self.root.attributes('-topmost', True)
        self.root.geometry(f"{self.root.winfo_screenwidth()}x{self.root.winfo_screenheight()}+0+0")
        self.root.configure(bg='black')

        self.canvas = tk.Canvas(self.root, cursor="cross", bg="black", highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)

        self.start_x = None
        self.start_y = None
        self.rect = None
        self.selection_box = None

        self.canvas.bind("<ButtonPress-1>", self.on_button_press)
        self.canvas.bind("<B1-Motion>", self.on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_button_release)
        
        self.root.bind('<Escape>', self.on_escape)
        self.root.focus_set()

    def on_escape(self, event):
        """Handler for Escape key to cancel selection."""
        logger.info("Selection cancelled by user (Escape).")
        self.selection_box = None
        self.root.quit()

    def on_button_press(self, event):
        self.start_x = event.x
        self.start_y = event.y
        if self.rect:
            self.canvas.delete(self.rect)
        self.rect = self.canvas.create_rectangle(
            self.start_x, self.start_y, self.start_x, self.start_y, 
            outline='red', width=2
        )

    def on_mouse_drag(self, event):
        self.canvas.coords(self.rect, self.start_x, self.start_y, event.x, event.y)

    def on_button_release(self, event):
        end_x = event.x
        end_y = event.y

        x1 = min(self.start_x, end_x)
        y1 = min(self.start_y, end_y)
        x2 = max(self.start_x, end_x)
        y2 = max(self.start_y, end_y)
        
        # Ensure a minimum size was selected
        if x2 - x1 > 10 and y2 - y1 > 10:
            self.selection_box = (int(x1), int(y1), int(x2), int(y2))
        
        self.root.quit()

    def get_selection(self):
        """Runs the window and returns the selected coordinates (x1, y1, x2, y2)."""
        try:
            self.root.mainloop()
            return self.selection_box
        except Exception as e:
            logger.error(f"Error in ScreenSelector: {e}")
            return None
        finally:
            try:
                self.root.destroy()
            except:
                pass
