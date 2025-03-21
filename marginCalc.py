import tkinter as tk
from tkinter import messagebox

class MarginCalculatorApp:
    def __init__(self, master):
        self.master = master
        master.title("Margin Calculator")

        tk.Label(master, text="Cost ($):").grid(row=0, column=0, sticky="e")
        tk.Label(master, text="Price ($):").grid(row=1, column=0, sticky="e")
        tk.Label(master, text="Kickback (%):").grid(row=2, column=0, sticky="e")
        tk.Label(master, text="Discount (%):").grid(row=3, column=0, sticky="e")

        self.cost_entry = tk.Entry(master)
        self.price_entry = tk.Entry(master)
        self.kickback_entry = tk.Entry(master)
        self.discount_entry = tk.Entry(master)

        self.cost_entry.grid(row=0, column=1, padx=10, pady=5)
        self.price_entry.grid(row=1, column=1, padx=10, pady=5)
        self.kickback_entry.grid(row=2, column=1, padx=10, pady=5)
        self.discount_entry.grid(row=3, column=1, padx=10, pady=5)

        self.result_label = tk.Label(master, text="", font=("Arial", 12), fg="green")
        self.result_label.grid(row=5, column=0, columnspan=2, pady=10)

        tk.Button(master, text="Calculate Margin", command=self.calculate_margin).grid(row=4, column=0, columnspan=2, pady=10)

    def calculate_margin(self):
        try:
            cost = float(self.cost_entry.get())
            price = float(self.price_entry.get())
            kickback_pct = float(self.kickback_entry.get()) / 100
            discount_pct = float(self.discount_entry.get()) / 100 if self.discount_entry.get() else 0

            discounted_price = price * (1 - discount_pct)
            kickback_value = cost * kickback_pct
            net_cost = cost - kickback_value

            if discounted_price == 0:
                raise ValueError("Discounted price cannot be zero.")

            margin = ((discounted_price - net_cost) / discounted_price) * 100

            self.result_label.config(text=f"Discounted Price: ${discounted_price:.2f}\n"
                                          f"Kickback Value: ${kickback_value:.2f}\n"
                                          f"Net Cost: ${net_cost:.2f}\n"
                                          f"Margin: {margin:.2f}%", fg="blue")
        except ValueError as e:
            messagebox.showerror("Input Error", f"Please enter valid numbers.\n\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = MarginCalculatorApp(root)
    root.mainloop()
