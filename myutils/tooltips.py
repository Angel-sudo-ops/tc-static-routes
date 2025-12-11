class ToolTip:
    def __init__(self, widget, text, delay=400, fade_duration=500):
        self.widget = widget
        self.text = text
        self.delay = delay  # delay before showing tooltip in milliseconds
        self.fade_duration = fade_duration  # duration of fade effect in milliseconds
        self.tooltip_window = None
        self.id = None
        self.opacity = 0
        self.is_fading_out = False

        self.widget.bind("<Enter>", self.schedule_tooltip)
        self.widget.bind("<Leave>", self.start_fade_out)
        self.widget.bind("<Button-1>", self.on_click)
        self.widget.winfo_toplevel().bind("<Motion>", self.check_motion)

    def schedule_tooltip(self, event):
        self.cancel_tooltip()
        if not self.is_fading_out:
            self.id = self.widget.after(self.delay, self.show_tooltip)

    def show_tooltip(self):
        if self.tooltip_window or not self.text:
            return
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 25

        self.tooltip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        tw.attributes('-alpha', 0.0)  # Start with full transparency

        style = ttk.Style()
        style.configure("Tooltip.TLabel", background="white", relief='solid', borderwidth=1, font=("helvetica", "8", "normal"))

        label = ttk.Label(tw, text=self.text, 
                          style="Tooltip.TLabel"
                        #   justify='left',
                        #   background="white", relief='solid', borderwidth=1,
                        #   font=("helvetica", "8", "normal")
                          )
        label.pack(ipadx=1)

        self.is_fading_out = False
        self.fade_in()

    def fade_in(self):
        if self.opacity < 1.0 and not self.is_fading_out:
            self.opacity += 0.05
            self.tooltip_window.attributes('-alpha', self.opacity)
            self.tooltip_window.after(int(self.fade_duration / 20), self.fade_in)
        else:
            if self.opacity >= 1.0:
                self.opacity = 1.0

    def start_fade_out(self, event=None):
        if self.tooltip_window and not self.is_fading_out:
            self.is_fading_out = True
            self.fade_out()

    def fade_out(self):
        if self.opacity > 0:
            self.opacity -= 0.05
            if self.tooltip_window:
                self.tooltip_window.attributes('-alpha', self.opacity)
                self.tooltip_window.after(int(self.fade_duration / 20), self.fade_out)
        else:
            if self.tooltip_window:
                self.tooltip_window.destroy()
                self.tooltip_window = None
                self.opacity = 0  # Reset opacity for the next tooltip
            self.is_fading_out = False

    def cancel_tooltip(self):
        if self.id:
            self.widget.after_cancel(self.id)
            self.id = None

    def check_motion(self, event):
        widget_under_cursor = self.widget.winfo_containing(event.x_root, event.y_root)
        if widget_under_cursor != self.widget and self.tooltip_window and not self.is_fading_out:
            self.start_fade_out()
    
    def on_click(self, event):
        # Reset the tooltip logic on click to ensure it can still appear
        self.cancel_tooltip()
        if self.tooltip_window:
            self.start_fade_out()
        self.schedule_tooltip(event)