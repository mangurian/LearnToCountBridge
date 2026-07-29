import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import copy

# Card definitions
SUITS = ['♠', '♥', '♦', '♣']
COLORS = {'♠': 'black', '♥': 'red', '♦': 'red', '♣': 'black'}
RANKS = ['A', 'K', 'Q', 'J', '10', '9', '8', '7', '6', '5', '4', '3', '2']
RANK_VALUES = {r: v for v, r in enumerate(reversed(RANKS), start=2)}
PLAYERS = ['N', 'E', 'S', 'W']
HCP_VALUES = {'A': 4, 'K': 3, 'Q': 2, 'J': 1}

class BridgeTracker(tk.Tk):
    def __init__(self):
        super().__init__()
        # Hide the main window while the splash screen is showing
        self.withdraw()
        
        self.title("Learn to Count")
        self.configure(bg="#2b2b2b")
        
        # State Variables
        self.trump_var = tk.StringVar(value='NT')
        self.status_var = tk.StringVar(value="Waiting for opening leader...")
        
        self.undo_stack = []
        self.current_turn = None
        self.opening_leader = None
        self.current_trick = []
        self.past_tricks = []
        self.played_cards = set()
        self.deck_widgets = {}
        
        self.show_splash()
        
    def show_splash(self):
        """Displays the initial splash screen."""
        splash = tk.Toplevel(self)
        splash.overrideredirect(True) # Removes window borders/controls
        splash.configure(bg="#0b5e2c") # Classic bridge table green
        splash.attributes("-topmost", True)
        
        # Center the splash screen
        splash_width = 500
        splash_height = 300
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width // 2) - (splash_width // 2)
        y = (screen_height // 2) - (splash_height // 2)
        splash.geometry(f"{splash_width}x{splash_height}+{x}+{y}")
        
        # Splash content
        tk.Label(splash, text="Learn to Count", font=("Arial", 36, "bold"), bg="#0b5e2c", fg="white").pack(expand=True)
        tk.Label(splash, text="copyright BGB 2026", font=("Arial", 12), bg="#0b5e2c", fg="lightgray").pack(side=tk.BOTTOM, pady=15)
        
        # Wait 3 seconds (3000 ms), then launch the main app
        self.after(3000, lambda: self.start_main_app(splash))
        
    def start_main_app(self, splash):
        """Destroys splash screen and launches the full-screen main UI."""
        splash.destroy()
        self.deiconify() # Bring back the main window
        
        self.build_ui()
        
        # Set initial fullscreen state
        self.set_fullscreen(True)
        
        # Keyboard bindings for Full Screen control
        self.bind("<Escape>", lambda e: self.set_fullscreen(False))
        self.bind("<F11>", self.toggle_fullscreen)
        
        # Prompt for leader shortly after UI builds
        self.after(500, self.prompt_for_leader)
        
    def build_ui(self):
        # --- Top Frame: The 52 Cards ---
        self.deck_frame = tk.Frame(self, bg="#2b2b2b")
        self.deck_frame.pack(side=tk.TOP, pady=20)
        
        tk.Label(self.deck_frame, text="(Press ESC to exit Full Screen)", bg="#2b2b2b", fg="gray", font=("Arial", 10)).grid(row=0, column=0, columnspan=13, pady=(0, 5))
        
        for i, suit in enumerate(SUITS):
            for j, rank in enumerate(RANKS):
                card_canvas = tk.Canvas(self.deck_frame, width=50, height=75, bg="#2b2b2b", highlightthickness=0)
                card_canvas.grid(row=i+1, column=j, padx=4, pady=4)
                
                card_canvas.bind("<Button-1>", lambda e, s=suit, r=rank: self.on_card_click(s, r))
                self.deck_widgets[(suit, rank)] = card_canvas
                
        # --- Middle Frame: Expanded Responsive Table ---
        self.middle_frame = tk.Frame(self, bg="#2b2b2b")
        self.middle_frame.pack(fill=tk.BOTH, expand=True, padx=20)
        
        self.table_canvas = tk.Canvas(self.middle_frame, bg="#0b5e2c", highlightthickness=2)
        self.table_canvas.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Bind the resize event to dynamically adjust the layout
        self.table_canvas.bind("<Configure>", lambda e: self.redraw())

        # --- Bottom Frame: Controls ---
        self.control_frame = tk.Frame(self, bg="#2b2b2b")
        self.control_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=10)
        
        tk.Label(self.control_frame, textvariable=self.status_var, font=("Arial", 16, "bold"), bg="#2b2b2b", fg="yellow").pack()
        
        btn_frame = tk.Frame(self.control_frame, bg="#2b2b2b")
        btn_frame.pack(fill=tk.X, padx=20, pady=10)
        
        # Left Side Buttons (Game Controls)
        left_btns = tk.Frame(btn_frame, bg="#2b2b2b")
        left_btns.pack(side=tk.LEFT)
        
        tk.Button(left_btns, text="Start New Hand", font=("Arial", 12, "bold"), command=self.new_hand, cursor="hand2").pack(side=tk.LEFT, padx=10)
        tk.Button(left_btns, text="Undo Last Click", font=("Arial", 12, "bold"), command=self.undo, cursor="hand2").pack(side=tk.LEFT, padx=10)
        
        tk.Label(left_btns, text="Trump Suit:", bg="#2b2b2b", fg="white", font=("Arial", 12, "bold")).pack(side=tk.LEFT, padx=10)
        trump_menu = ttk.Combobox(left_btns, textvariable=self.trump_var, values=['NT', '♠', '♥', '♦', '♣'], state="readonly", width=5, font=("Arial", 14))
        trump_menu.pack(side=tk.LEFT)

        # Right Side Buttons (App Controls)
        right_btns = tk.Frame(btn_frame, bg="#2b2b2b")
        right_btns.pack(side=tk.RIGHT)
        
        tk.Button(right_btns, text="Help", font=("Arial", 12, "bold"), command=self.show_help, cursor="hand2", bg="#3a7ca5", fg="white").pack(side=tk.LEFT, padx=10)
        
        self.fs_btn = tk.Button(right_btns, text="Exit Full Screen", font=("Arial", 12, "bold"), command=self.toggle_fullscreen, cursor="hand2")
        self.fs_btn.pack(side=tk.LEFT, padx=10)

    # --- App Control Methods ---
    def set_fullscreen(self, state):
        self.attributes("-fullscreen", state)
        if hasattr(self, 'fs_btn'):
            if state:
                self.fs_btn.config(text="Exit Full Screen")
            else:
                self.fs_btn.config(text="Enter Full Screen")

    def toggle_fullscreen(self, event=None):
        current_state = self.attributes("-fullscreen")
        self.set_fullscreen(not current_state)

    def show_help(self):
        """Displays a pop-up window explaining how to use the application."""
        help_win = tk.Toplevel(self)
        help_win.title("How to Use Learn to Count")
        help_win.geometry("550x530")
        help_win.configure(bg="#2b2b2b")
        help_win.transient(self)
        
        title_lbl = tk.Label(help_win, text="Learn to Count", font=("Arial", 18, "bold"), bg="#2b2b2b", fg="white")
        title_lbl.pack(pady=(20, 10))
        
        instructions = (
            "1. Starting a Hand: When you open the app or click 'Start New Hand', "
            "you will be prompted to select who is making the Opening Lead (N, E, S, or W).\n\n"
            "2. Playing Cards: Click on a card in the 52-card deck at the top to play it. "
            "It will automatically jump to the table in front of the active player.\n\n"
            "3. Game Flow: Play proceeds clockwise. After 4 cards are played, the program "
            "determines who won the trick based on the Trump Suit (set at the bottom). "
            "The winner of the trick automatically leads the next trick.\n\n"
            "4. Hand History: As cards are played, they cascade cleanly in front of each "
            "player on the table. You can visually track the entire history of the hand.\n\n"
            "5. HCP Tracking: The High Card Points (A=4, K=3, Q=2, J=1) played by each player "
            "are calculated automatically and displayed under their name.\n\n"
            "6. Undo: Made a mistake? Click 'Undo Last Click' at any time to step backwards."
        )
        
        text_lbl = tk.Label(help_win, text=instructions, font=("Arial", 12), bg="#2b2b2b", fg="lightgray", justify=tk.LEFT, wraplength=500)
        text_lbl.pack(padx=20, pady=5)
        
        # Red Warning Caveat
        caveat_text = "WARNING: This application should not be used during competition. It is solely designed to facilitate learning to count on practice hands."
        warning_lbl = tk.Label(help_win, text=caveat_text, font=("Arial", 11, "bold"), bg="#2b2b2b", fg="#ff6666", justify=tk.LEFT, wraplength=500)
        warning_lbl.pack(padx=20, pady=(15, 5))
        
        tk.Button(help_win, text="Close Help", font=("Arial", 12, "bold"), command=help_win.destroy, cursor="hand2").pack(pady=15)
        
        # Center the window
        help_win.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() // 2) - (help_win.winfo_width() // 2)
        y = self.winfo_y() + (self.winfo_height() // 2) - (help_win.winfo_height() // 2)
        help_win.geometry(f"+{x}+{y}")

    # --- Drawing Methods ---
    def draw_realistic_card(self, canvas, x, y, suit, rank, width, height, tags=""):
        """Procedurally draws a realistic looking playing card on a given canvas."""
        color = COLORS[suit]
        
        canvas.create_rectangle(x, y, x+width, y+height, fill="white", outline="black", width=2, tags=tags)
        
        small_font = max(8, int(width * 0.2))
        large_font = max(16, int(width * 0.45))
        
        canvas.create_text(x + 12, y + 12, text=rank, font=("Arial", small_font, "bold"), fill=color, tags=tags)
        canvas.create_text(x + 12, y + 12 + small_font, text=suit, font=("Arial", small_font), fill=color, tags=tags)
        
        canvas.create_text(x + width - 12, y + height - 12 - small_font, text=rank, font=("Arial", small_font, "bold"), fill=color, tags=tags)
        canvas.create_text(x + width - 12, y + height - 12, text=suit, font=("Arial", small_font), fill=color, tags=tags)
        
        canvas.create_text(x + width/2, y + height/2, text=suit, font=("Arial", large_font), fill=color, tags=tags)

    def redraw(self):
        """Redraws the canvas UI using the realistic card drawing method."""
        for card, canvas in self.deck_widgets.items():
            canvas.delete("all")
            if card in self.played_cards:
                canvas.config(cursor="arrow")
            else:
                suit, rank = card
                self.draw_realistic_card(canvas, 2, 2, suit, rank, width=46, height=71)
                canvas.config(cursor="hand2")
                
        w = self.table_canvas.winfo_width()
        h = self.table_canvas.winfo_height()
        
        if w < 100 or h < 100:
            return
            
        self.table_canvas.delete("all")
        
        all_played = []
        for trick in self.past_tricks:
            all_played.extend(trick)
        all_played.extend(self.current_trick)
        
        player_hcp = {p: 0 for p in PLAYERS}
        player_card_counts = {p: 0 for p in PLAYERS}
        
        for p, (suit, rank) in all_played:
            player_hcp[p] += HCP_VALUES.get(rank, 0)
            
        player_coords = {
            'N': (w // 2, 55),
            'E': (w - 75, h // 2),
            'S': (w // 2, h - 55),
            'W': (75, h // 2)
        }
        
        for p, (x, y) in player_coords.items():
            fill_color = "yellow" if self.current_turn == p else "lightgray"
            box_id = self.table_canvas.create_rectangle(x-45, y-40, x+45, y+40, fill=fill_color, tags=f"player_{p}")
            text_id = self.table_canvas.create_text(x, y, text=f"{p}\n{player_hcp[p]} HCP", font=("Arial", 14, "bold"), justify=tk.CENTER, tags=f"player_text_{p}")
            self.table_canvas.tag_bind(box_id, "<Button-1>", lambda e, pl=p: self.on_player_click(pl))
            self.table_canvas.tag_bind(text_id, "<Button-1>", lambda e, pl=p: self.on_player_click(pl))

        for p, (suit, rank) in all_played:
            idx = player_card_counts[p]
            c_width, c_height = 80, 120
            
            if p == 'N':
                cx = (w // 2) - 220 + (idx * 35)
                cy = 130
            elif p == 'S':
                cx = (w // 2) - 220 + (idx * 35)
                cy = h - 250
            elif p == 'W':
                cx = 150
                cy = (h // 2) - 280 + (idx * 40)
            elif p == 'E':
                cx = w - 230
                cy = (h // 2) - 280 + (idx * 40)
                
            self.draw_realistic_card(self.table_canvas, cx, cy, suit, rank, width=c_width, height=c_height, tags="card")
            player_card_counts[p] += 1

    # --- Game Logic Methods ---
    def prompt_for_leader(self):
        prompt = tk.Toplevel(self)
        prompt.title("Opening Leader")
        prompt.geometry("400x180")
        prompt.configure(bg="#2b2b2b")
        prompt.transient(self) 
        prompt.grab_set() 
        prompt.protocol("WM_DELETE_WINDOW", lambda: None)
        
        tk.Label(prompt, text="Who is the opening leader?", font=("Arial", 16, "bold"), bg="#2b2b2b", fg="white").pack(pady=30)
        
        btn_frame = tk.Frame(prompt, bg="#2b2b2b")
        btn_frame.pack()
        
        def select(p):
            self.save_state()
            self.opening_leader = p
            self.current_turn = p
            self.status_var.set(f"Opening leader selected. Waiting for {p} to play.")
            self.redraw()
            prompt.destroy()
            
        for p in PLAYERS:
            tk.Button(btn_frame, text=p, font=("Arial", 14, "bold"), width=4, cursor="hand2", command=lambda pl=p: select(pl)).pack(side=tk.LEFT, padx=15)

        prompt.update_idletasks()
        w, h = prompt.winfo_width(), prompt.winfo_height()
        x = self.winfo_x() + (self.winfo_width() // 2) - (w // 2)
        y = self.winfo_y() + (self.winfo_height() // 2) - (h // 2)
        prompt.geometry(f"+{x}+{y}")

    def on_player_click(self, player):
        if len(self.played_cards) > 0 or self.current_turn is not None: return  
        self.save_state()
        self.opening_leader, self.current_turn = player, player
        self.status_var.set(f"Opening leader selected. Waiting for {player} to play.")
        self.redraw()
        
    def on_card_click(self, suit, rank):
        card = (suit, rank)
        if card in self.played_cards: return 
        if self.current_turn is None:
            messagebox.showwarning("Select Leader", "Please select an opening leader first.")
            return
            
        self.save_state()
        
        if len(self.current_trick) == 4:
            self.past_tricks.append(list(self.current_trick))
            self.current_trick = []
        
        self.played_cards.add(card)
        self.current_trick.append((self.current_turn, card))
        
        if len(self.current_trick) < 4:
            self.current_turn = PLAYERS[(PLAYERS.index(self.current_turn) + 1) % 4]
            self.status_var.set(f"Waiting for {self.current_turn} to play.")
        else:
            winner = self.determine_winner(self.current_trick)
            self.current_turn = winner
            self.status_var.set(f"Trick won by {winner}. Waiting for {winner} to lead next trick.")
            
        self.redraw()
        
    def determine_winner(self, trick):
        lead_suit = trick[0][1][0]
        winner, highest_val = trick[0][0], RANK_VALUES[trick[0][1][1]]
        trump_played = (lead_suit == self.trump_var.get())
        
        for p, (suit, rank) in trick[1:]:
            val, is_trump = RANK_VALUES[rank], (suit == self.trump_var.get())
            if is_trump:
                if not trump_played:
                    trump_played, winner, highest_val = True, p, val
                elif val > highest_val:
                    winner, highest_val = p, val
            elif suit == lead_suit and not trump_played:
                if val > highest_val:
                    winner, highest_val = p, val
        return winner
        
    def save_state(self):
        state = {
            'turn': self.current_turn, 'leader': self.opening_leader,
            'trick': copy.deepcopy(self.current_trick), 'past': copy.deepcopy(self.past_tricks),
            'played': set(self.played_cards)
        }
        self.undo_stack.append(state)
        
    def undo(self):
        if not self.undo_stack: return
        state = self.undo_stack.pop()
        self.current_turn, self.opening_leader = state['turn'], state['leader']
        self.current_trick, self.past_tricks = state['trick'], state['past']
        self.played_cards = state['played']
        
        if self.current_turn is None: self.status_var.set("Waiting for opening leader...")
        elif len(self.current_trick) == 4: self.status_var.set(f"Trick won by {self.current_turn}. Waiting for {self.current_turn} to lead next trick.")
        else: self.status_var.set(f"Waiting for {self.current_turn} to play.")
        self.redraw()
        
    def new_hand(self):
        self.save_state() 
        self.current_turn, self.opening_leader = None, None
        self.current_trick, self.past_tricks, self.played_cards = [], [], set()
        self.status_var.set("Waiting for opening leader...")
        self.redraw()
        self.prompt_for_leader()

if __name__ == "__main__":
    app = BridgeTracker()
    app.mainloop()