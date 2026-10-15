#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# ==========================================================================
#  Bridge Tracker   --   version 1.7            (2026-10-13)
#
#  What changed in v1.7 (this revision):
#
#    * the winner of a trick leads the next one BY ITSELF -- there is
#      nothing to click any more.  The app works the winner out (highest
#      trump, or the highest card of the led suit in no-trump), names it
#      in the status line and the card window, and puts the turn straight
#      on that player's box.  The "Winner leads next trick" tick has gone,
#      because automatic leading is now the only way the app works;
#    * the cards a player has played are printed in the colour of their
#      suit: red for hearts and diamonds, black for clubs and spades.  An
#      honour keeps its tinted cell, so the points are still easy to see;
#    * the note in the corner of a player's box -- "on turn" or "leads
#      trick N" -- FLASHES, so it is obvious at a glance whose action it
#      is.  The orange edge round that player's panel stays as it was.
#
#  (v1.6, for reference: the play went clockwise one player at a time and
#  the user clicked the box of the player who was to lead each trick, with
#  a "Winner leads next trick" tick as the alternative; a card that had
#  already been played could not be clicked and the card window never
#  recorded a card in a suit the user had not picked; every player's box
#  was labelled "Cards Played"; the window label carried "@BGB 2026".)
#
#  (v1.5, for reference: the version number went into the app name and the
#  window title, the board was drawn 70% bigger, bidding and play became
#  clockwise with no direction to choose, each player got an "estimated
#  points" box and a "points played" box, helper windows started at the
#  upper right of the board, the suit led was pre-selected in the card
#  window, and the trick winner was worked out and led the next trick.)
# ==========================================================================
"""
Bridge Tracker  --  a point-and-click Python replacement for the
"Bridge Tracker.xlsm" spreadsheet.

What it does (mirrors the spreadsheet, plus the extras that were asked for):

  * Four hand panels (North / East / South / West) exactly where the
    spreadsheet puts them, colour coded the same way (North light blue,
    East light green, South light orange, West light gold).
  * Each panel has one row per suit (C, D, H, S) with ten card cells, and
    two point boxes underneath: "estimated points" and "points played".
  * The bidding goes in the middle of the board: one column per player, one
    row per round of the auction, and the calls always run clockwise.
  * Start of a hand: a small pop-up asks who deals.  The auction starts with
    the dealer and moves clockwise.
  * The bid pop-up only ever contains point-and-click buttons (levels 1-7,
    the five strains, Pass / Double / Redouble, and "Bidding is over").
    Every bid is written straight into the bidding grid, then the pop-up
    moves on to the next bidder.
  * When the bidding ends the pop-up turns into a "card played" pop-up.
    The opening lead belongs to the player to the left of the highest
    bidder; play runs clockwise from there, and every card lands in the
    right suit row of the right player.
  * The suit that was led is the default suit to click, so a player who
    follows suit is one click; a different suit can be chosen for that
    card, and the next player starts again on the led suit.  A card that
    somebody has already played cannot be clicked.
  * The winner of a trick (trump or no-trump) is found automatically,
    named, and given the lead of the next trick -- there is no click to
    make.  The cards in each player's panel are printed in the colour of
    their suit (red for hearts and diamonds, black for clubs and spades),
    and the note that says who is on turn, or who leads the trick,
    flashes so that the player whose action it is can be seen at a glance.
  * A, K, Q and J add 4, 3, 2 and 1 point to that player's "points played"
    box.  The "estimated points" box is the user's own pre-play estimate and
    never changes once it is set.
  * Undo, a "Clear Entries" button (the spreadsheet's macro) and a small
    on-board help window.

Only the Python standard library is used (tkinter).  Nothing is typed at
the keyboard in normal use: every choice is a button click.

Run:      python bridge_tracker.py          (or double-click the .pyw)
Demo:     python bridge_tracker.py --demo   (scripted two tricks, then free play)
"""

from __future__ import annotations

import copy
import datetime
import sys
import tkinter as tk
from tkinter import messagebox

__version__ = "1.7 (2026-10-13)"

# --------------------------------------------------------------------------
#  Bridge vocabulary / model
# --------------------------------------------------------------------------

SEATS = ("N", "E", "S", "W")                     # clockwise order
SEAT_NAME = {"N": "North", "E": "East", "S": "South", "W": "West"}

SUITS = ("C", "D", "H", "S")
SUIT_LETTER = {"C": "C", "D": "D", "H": "H", "S": "S"}
SUIT_SYMBOL = {"C": "\u2663", "D": "\u2666", "H": "\u2665", "S": "\u2660"}
SUIT_COLOR = {"C": "#1a1a1a", "D": "#C00000", "H": "#C00000", "S": "#1a1a1a"}
SUIT_WORD = {"C": "Clubs", "D": "Diamonds", "H": "Hearts", "S": "Spades"}

RANKS = ("A", "K", "Q", "J", "10", "9", "8", "7", "6", "5", "4", "3", "2")
HONOR_POINTS = {"A": 4, "K": 3, "Q": 2, "J": 1}

STRAINS = ("C", "D", "H", "S", "NT")
STRAIN_NAME = {"C": "Clubs", "D": "Diamonds", "H": "Hearts",
               "S": "Spades", "NT": "No-trump"}
STRAIN_SYMBOL = {"C": "\u2663", "D": "\u2666", "H": "\u2665",
                 "S": "\u2660", "NT": "NT"}

BIDS = tuple("%d%s" % (lvl, st) for lvl in range(1, 8) for st in STRAINS)
CALLS = ("P", "X", "XX") + BIDS

CALL_TEXT = {"P": "Pass", "X": "Dbl", "XX": "Rdbl"}


def bid_text(call):
    """'1H' -> '1\u2665', 'P' -> 'Pass', ..."""
    if call in CALL_TEXT:
        return CALL_TEXT[call]
    return call[0] + STRAIN_SYMBOL[call[1:]]


def bid_value(call):
    """Sort key for a level/strain bid."""
    return (int(call[0]), STRAINS.index(call[1:]))


def step(seat):
    """The player who sits next to `seat` -- always clockwise (v1.5)."""
    return SEATS[(SEATS.index(seat) + 1) % 4]


def left_of(seat):
    """Left-hand opponent: the player to the declarer's left
    (= the next player clockwise)."""
    return step(seat)


def side_of(seat):
    """0 = North/South partnership, 1 = East/West partnership."""
    return SEATS.index(seat) % 2


def seat_letter_side(seat):
    return "N-S" if side_of(seat) == 0 else "E-W"


def card_name(suit, rank):
    return "%s%s" % (rank, SUIT_SYMBOL[suit])


class Auction(object):
    """The bidding: who has called what, and whose turn it is.

    v1.5: the auction is always clockwise -- there is no direction choice.
    """

    def __init__(self, dealer):
        self.dealer = dealer
        self.calls = []                      # [(seat, call), ...] in order
        self.turn = dealer

    # -- ordering helpers ---------------------------------------------------
    def seat_order(self):
        """The four seats in the order in which they call: dealer first."""
        seq, s = [self.dealer], self.dealer
        for _ in range(3):
            s = step(s)
            seq.append(s)
        return seq

    def column_of(self, seat):
        return self.seat_order().index(seat)

    def last_level_bid(self):
        """(index, seat, call) of the last level/strain bid, or None."""
        for i in range(len(self.calls) - 1, -1, -1):
            seat, call = self.calls[i]
            if call in BIDS:
                return i, seat, call
        return None

    def last_call(self):
        return self.calls[-1][1] if self.calls else None

    # -- legality -----------------------------------------------------------
    def is_legal(self, call):
        seat = self.turn
        last = self.last_level_bid()
        if call == "P":
            return True
        if call == "X":
            if last is None:
                return False
            idx, lseat, _ = last
            if side_of(lseat) == side_of(seat):
                return False                      # cannot double partner
            return not any(c == "X" for _, c in self.calls[idx + 1:])
        if call == "XX":
            if last is None:
                return False
            idx, lseat, _ = last
            after = self.calls[idx + 1:]
            doubler = None
            for s, c in after:
                if c == "X":
                    doubler = s
            if doubler is None or any(c == "XX" for _, c in after):
                return False
            return side_of(doubler) != side_of(seat)
        if call in BIDS:
            if last is None:
                return True
            return bid_value(call) > bid_value(last[2])
        return False

    def illegal_reason(self, call):
        if call == "P":
            return ""
        last = self.last_level_bid()
        if call == "X":
            return "no bid to double" if last is None else "double not available"
        if call == "XX":
            return "redouble not available"
        if last is not None and bid_value(call) <= bid_value(last[2]):
            return "must be higher than %s" % bid_text(last[2])
        return ""

    def place(self, call):
        seat = self.turn
        self.calls.append((seat, call))
        self.turn = step(seat)                # clockwise, always
        return seat

    # -- auction outcome ----------------------------------------------------
    def passed_out(self):
        return self.last_level_bid() is None and len(self.calls) >= 4

    def contract(self):
        """{'level','strain','doubled','declarer','bidder'} or None."""
        last = self.last_level_bid()
        if last is None:
            return None
        idx, seat, call = last
        level, strain = int(call[0]), call[1:]
        doubled = 0
        for _, c in self.calls[idx + 1:]:
            if c == "XX":
                doubled = 2
            elif c == "X" and not doubled:
                doubled = 1
        # declarer = first player of the winning side who named that strain
        declarer = seat
        for s, c in self.calls:
            if c in BIDS and c[1:] == strain and side_of(s) == side_of(seat):
                declarer = s
                break
        return {"level": level, "strain": strain, "doubled": doubled,
                "declarer": declarer, "bidder": seat}

    def auction_text(self):
        return "  ".join("%s %s" % (s, bid_text(c)) for s, c in self.calls)


def contract_text(contract):
    if contract is None:
        return "\u2014"
    t = "%d%s" % (contract["level"], STRAIN_SYMBOL[contract["strain"]])
    if contract["doubled"] == 1:
        t += " x"
    elif contract["doubled"] == 2:
        t += " xx"
    return "%s by %s" % (t, SEAT_NAME[contract["declarer"]])


def trump_of(contract):
    """The trump suit of a contract, or None for a no-trump contract."""
    if contract is None:
        return None
    strain = contract["strain"]
    return strain if strain in SUITS else None


def beats(card, best, trump):
    """Does the card `card` beat `best` in the trick that is being played?

    `card`, `best` are (seat, suit, rank).  The first card of the trick is
    always of the led suit, so anything that is not the led suit can only
    win by being a trump.
    """
    cs, cr = card[1], card[2]
    bs, br = best[1], best[2]
    if cs == bs:
        return RANKS.index(cr) < RANKS.index(br)     # RANKS runs high -> low
    if trump is None:
        return False                                 # no trump: led suit wins
    if cs == trump:
        return True
    if bs == trump:
        return False
    return False


def win_reason(record):
    """Why that card won the trick -- shown to the user."""
    card = card_name(record["card"][0], record["card"][1])
    if record["trump"]:
        return "%s, a trump" % card
    return "%s, the highest %s" % (card, SUIT_WORD[record["led"]])


class Play(object):
    """The cards played: one list per suit inside each player's panel.

    v1.7: play is always clockwise, one player at a time, and the winner of
    a trick (highest trump, or the highest card of the led suit when the
    contract is no-trump) leads the next trick by itself.  There is nothing
    to click between tricks: when the fourth card of a trick lands, the
    winner is worked out, recorded, and the turn is put straight on that
    player's box.
    """

    def __init__(self, leader, trump=None):
        self.leader = leader                  # opening leader (declarer's LHO)
        self.trump = trump if trump in SUITS else None    # None == no trump
        self.turn = leader
        self.cards = dict((s, dict((su, []) for su in SUITS)) for s in SEATS)
        self.sequence = []                    # [(seat, suit, rank), ...]
        self.trick_log = []                   # one record per completed trick

    # -- placing cards ------------------------------------------------------
    def place(self, suit, rank):
        if self.complete():
            return None                       # the hand is over
        seat = self.turn
        self.cards[seat][suit].append(rank)
        self.sequence.append((seat, suit, rank))
        if len(self.sequence) % 4 == 0:
            self._close_trick()               # four cards: find the winner
        else:
            self.turn = step(seat)
        return seat

    def _close_trick(self):
        cards = self.sequence[-4:]
        seat, card, trumped = self.trick_winner(cards)
        self.trick_log.append({"n": len(self.trick_log) + 1,
                               "leader": cards[0][0], "led": cards[0][1],
                               "winner": seat, "card": card,
                               "trump": trumped})
        # v1.7: the winner of the trick takes the lead of the next one.  No
        # click is needed anywhere: the turn lands on the winner's box.
        self.turn = seat

    def trick_winner(self, cards):
        """(seat, (suit, rank), won_with_trump) for one completed trick."""
        best = cards[0]
        for card in cards[1:]:
            if beats(card, best, self.trump):
                best = card
        return best[0], (best[1], best[2]), bool(self.trump and best[1] == self.trump)

    # -- the trick in progress ---------------------------------------------
    def trick_cards(self):
        """The cards of the trick being played (up to four, in order)."""
        n = len(self.sequence) % 4
        return self.sequence[-n:] if n else []

    def led_suit(self):
        """The suit led in the current trick, or None before the first card."""
        tc = self.trick_cards()
        return tc[0][1] if tc else None

    def last_record(self):
        return self.trick_log[-1] if self.trick_log else None

    # -- who leads the next trick ------------------------------------------
    def at_trick_start(self):
        """True at the start of a trick, before its first card has been
        played (and false once the hand is complete)."""
        return not self.complete() and len(self.sequence) % 4 == 0

    def leader_from_last_trick(self):
        """The seat the lead of the current trick came from: the winner of
        the trick before it (v1.7), or None at the start of a hand."""
        rec = self.last_record()
        if rec is None or not self.at_trick_start():
            return None
        return rec["winner"]

    # -- scoring / counters -------------------------------------------------
    def played(self):
        return set((su, r) for _, su, r in self.sequence)

    def honors(self, seat):
        return sum(HONOR_POINTS[r] for s, su, r in self.sequence
                   if s == seat and r in HONOR_POINTS)

    def honor_list(self, seat):
        return [(su, r) for s, su, r in self.sequence
                if s == seat and r in HONOR_POINTS]

    def tricks_won(self, seat):
        return sum(1 for rec in self.trick_log if rec["winner"] == seat)

    def side_tricks(self):
        """{'N-S': n, 'E-W': n} from the tricks completed so far."""
        out = {"N-S": 0, "E-W": 0}
        for rec in self.trick_log:
            out[seat_letter_side(rec["winner"])] += 1
        return out

    def count(self):
        return len(self.sequence)

    def complete(self):
        return len(self.sequence) >= 52

    def trick_no(self):
        """The trick that is being played now (1-13)."""
        return len(self.sequence) // 4 + 1 if len(self.sequence) < 52 else 13

    def card_in_trick(self):
        return len(self.sequence) % 4 + 1


# --------------------------------------------------------------------------
#  Board geometry -- one place that decides where everything sits on screen.
#  v1.5: every measurement below is a v1.0 measurement put through px(), so
#  the whole board can be resized from the single UI_SCALE number.
# --------------------------------------------------------------------------

UI_SCALE = 1.7         # v1.5: 70% bigger than the v1.0 board (1100 -> 1817 px)


def px(n):
    """A v1.0 pixel measurement, scaled for the v1.5 board."""
    return int(round(n * UI_SCALE))


def font_px(n):
    """A v1.0 font size, scaled (negative = Tk's pixel-sized fonts)."""
    return -max(8, int(round(abs(n) * UI_SCALE)))


POPUP_SCALE = 1.4      # the helper windows are scaled a little less than the
#                        board so that they still fit the empty corner


def pfont_px(n):
    """A v1.0 font size for a helper window, scaled for v1.5."""
    return -max(8, int(round(abs(n) * POPUP_SCALE)))


SLOT_W = px(30)        # width of one card cell
CELL_H = px(26)        # height of one card cell / suit row
SUIT_LABEL_W = px(30)  # width of the "C / D / H / S" label column
PANEL_PAD = px(6)
NAME_H = px(24)
EST_LABEL_H = px(20)
EST_VALUE_H = px(28)
POINTS_GAP = px(8)     # gap between the "estimated" and "played" boxes
SLOTS = 10             # card cells per suit row (same as the spreadsheet)

BID_HEADER_H = px(34)  # room for the "BIDDING" title line + the seat names
BID_PITCH = px(21)     # nominal height of one round of bidding
BID_ROWS_MIN = 10      # rounds drawn when the auction is still short

MARGIN = px(10)
GAP = px(12)

PANEL_W = PANEL_PAD * 2 + SUIT_LABEL_W + SLOTS * SLOT_W          # 581
PANEL_H = (PANEL_PAD * 2 + NAME_H + 4 * CELL_H
           + EST_LABEL_H + EST_VALUE_H)                          # 319
BID_W = PANEL_W
BID_H = BID_HEADER_H + BID_ROWS_MIN * BID_PITCH                  # 401

BOARD_W = MARGIN * 2 + PANEL_W * 3 + GAP * 2                     # 1817
BOARD_H = MARGIN * 2 + PANEL_H * 2 + BID_H + GAP * 2             # 1113

CENTER_X = MARGIN + PANEL_W + GAP
EAST_X = MARGIN + 2 * (PANEL_W + GAP)
WEST_TOP = (BOARD_H - PANEL_H) // 2

PANEL_RECT = {
    "N": (CENTER_X, MARGIN, PANEL_W, PANEL_H),
    "S": (CENTER_X, MARGIN + PANEL_H + GAP + BID_H + GAP, PANEL_W, PANEL_H),
    "W": (MARGIN, WEST_TOP, PANEL_W, PANEL_H),
    "E": (EAST_X, WEST_TOP, PANEL_W, PANEL_H),
}
BID_RECT = (CENTER_X, MARGIN + PANEL_H + GAP, BID_W, BID_H)

# The four empty areas around the table -- helper windows start in one of
# these so that nothing on the board is ever hidden.
CORNER_RECT = {
    "N": (0, 0, CENTER_X, WEST_TOP),                                   # top left
    "E": (EAST_X, 0, PANEL_W + MARGIN, WEST_TOP),                      # top right
    "W": (0, WEST_TOP + PANEL_H, CENTER_X, BOARD_H - (WEST_TOP + PANEL_H)),
    "S": (EAST_X, WEST_TOP + PANEL_H, PANEL_W + MARGIN,
          BOARD_H - (WEST_TOP + PANEL_H)),
}

# v1.5: every helper window opens here (the upper right corner of the board)
# and is then left exactly where the user drags it.
POPUP_START_CORNER = "E"

HINT_WRAP = px(280)    # wrap width for the board's own hint lines
POPUP_WRAP = int(round(300 * POPUP_SCALE))     # ... and for the helper windows


def content_rects():
    """Every rectangle on the board that carries information."""
    rects = []
    for seat, (x, y, w, h) in PANEL_RECT.items():
        rects.append((seat, x, y, x + w, y + h))
    x, y, w, h = BID_RECT
    rects.append(("bidding", x, y, x + w, y + h))
    return rects


def fits_in_corner(width, height, seat):
    """True when a window of this size can sit inside `seat`'s corner."""
    x, y, w, h = CORNER_RECT[seat]
    return width <= w - 16 and height <= h - 42


def seat_at(x, y):
    """Which player's panel holds this canvas coordinate (or None)."""
    for seat in SEATS:
        px_, py_, pw, ph = PANEL_RECT[seat]
        if px_ <= x <= px_ + pw and py_ <= y <= py_ + ph:
            return seat
    return None


# --------------------------------------------------------------------------
#  Colours / fonts
# --------------------------------------------------------------------------

PANEL_FILL = {"N": "#DEEBF7", "E": "#E2EFDA", "S": "#FBE5D6", "W": "#FFF2CC"}
PANEL_EDGE = "#7F7F7F"
CELL_FILL = "#FFFFFF"
CELL_EDGE = "#BFBFBF"
HONOR_FILL = "#FFF2CC"        # a card that scored points
HONOR_TEXT = "#B00020"        # ... the colour of the "points played" digits
CARD_INK = {"C": "#000000", "D": "#C00000", "H": "#C00000", "S": "#000000"}
#                             # v1.7: a card in a player's panel is printed in
#                             # the colour of its own suit -- red for hearts
#                             # and diamonds, black for clubs and spades
NOTE_TEXT = "#843C0C"         # the small note in a player's name row
FLASH_TEXT = "#C00000"        # v1.7: ... and the colour that note flashes in
BLINK_MS = 520                # v1.7: half the period of the flash, in ms
ACTIVE_EDGE = "#ED7D31"       # the player who is on turn
LED_FILL = "#DEEBF7"          # the suit that was led (the default suit)
PICK_FILL = "#FFE699"         # the suit / seat the user picked
BID_GRID_EDGE = "#8EA9DB"
BOARD_BG = "#F5F5F5"
HEADER_BG = "#44546A"
STATUS_BG = "#EDEDED"
BUTTON_BG = "#E7E6E6"

F_TITLE = ("Segoe UI", font_px(15), "bold")
F_NAME = ("Segoe UI", font_px(15), "bold")
F_SUIT = ("Segoe UI", font_px(14), "bold")
F_CARD = ("Segoe UI", font_px(14), "bold")
F_SMALL = ("Segoe UI", font_px(11))
F_TINY = ("Segoe UI", font_px(10))
# v1.7: the note that flashes when it is a player's turn (or their lead)
F_NOTE = ("Segoe UI", font_px(11), "bold")
F_BIG = ("Segoe UI", font_px(19), "bold")
F_BID = ("Segoe UI", font_px(13), "bold")
F_BTN = ("Segoe UI", font_px(12))
F_BTN_BIG = ("Segoe UI", font_px(13))
F_BTN_HEAD = ("Segoe UI", font_px(11), "bold")
F_HEAD = ("Segoe UI", font_px(13), "bold")
F_BOLD = ("Segoe UI", font_px(13), "bold")

# The helper windows use slightly smaller type than the board: they have to
# stay inside the empty corner (see POPUP_SCALE).
F_P_HEAD = ("Segoe UI", pfont_px(13), "bold")
F_P_SMALL = ("Segoe UI", pfont_px(11))
F_P_TINY = ("Segoe UI", pfont_px(10))
F_P_BTN = ("Segoe UI", pfont_px(12))
F_P_BTN_BIG = ("Segoe UI", pfont_px(13))
F_P_BTN_HEAD = ("Segoe UI", pfont_px(11), "bold")

APP_NAME = "CONTRACT BRIDGE TRACKER v1.7"
CREDIT = "@BGB 2026"      # v1.6: small credit line in the window label

FLASH_NOTES = ("on turn", "leads trick")


def note_flashes(note):
    """v1.7: the two notes that blink -- the ones that say whose action it
    is (whose turn it is, or who leads the trick)."""
    return note.startswith(FLASH_NOTES)


# --------------------------------------------------------------------------
#  The application
# --------------------------------------------------------------------------

class BridgeTracker(tk.Tk):

    def __init__(self, demo=False):
        tk.Tk.__init__(self)
        self.title("Bridge Tracker v1.7  \u2014  Contract Bridge hand recorder"
                   "  \u2014  @BGB 2026")
        self.configure(bg=BOARD_BG)
        self.resizable(False, False)

        self.state = None
        self.snapshots = []
        self.popup = None
        self.help_win = None
        self.demo = demo
        # v1.7: the turn follows the cards -- the winner of a trick leads the
        # next one -- so there is no "who leads?" question and no tick to
        # switch the automatic leader on and off any more.
        self.blink_items = []          # canvas ids of the notes that flash
        self.blink_on = True           # ... shown or hidden right now
        self.blink_job = None          # the pending after() callback

        self._build_toolbar()
        self._build_board()
        self._build_status()

        self.update_idletasks()
        self._center_window()

        self.new_hand()
        self.start_blinking()
        if demo:
            self.after(400, self.run_demo)

    # ------------------------------------------------------------------ UI
    def _build_toolbar(self):
        bar = tk.Frame(self, bg=HEADER_BG)
        bar.pack(side="top", fill="x")
        self.toolbar = bar

        self.date_var = tk.StringVar(value=datetime.date.today().strftime("%B %d, %Y"))
        tk.Label(bar, textvariable=self.date_var, bg=HEADER_BG, fg="white",
                 font=F_SMALL).pack(side="left", padx=(10, 8), pady=8)

        tk.Label(bar, text=APP_NAME, bg=HEADER_BG, fg="white",
                 font=F_TITLE).pack(side="left", padx=12, pady=8)

        # v1.6: the small credit line that belongs to the window label
        tk.Label(bar, text=CREDIT, bg=HEADER_BG, fg="#F2F2F2",
                 font=F_TINY).pack(side="left", padx=(0, 10), pady=(14, 2))

        self.toolbar_buttons = []
        for text, cmd in (("Quit", self.destroy),
                          ("Help", self.show_help),
                          ("Set Estimates", self.show_estimates),
                          ("Clear Entries", self.clear_entries),
                          ("Undo", self.undo),
                          ("New Hand", self.new_hand)):
            b = tk.Button(bar, text=text, command=cmd, font=F_BTN,
                          bg=BUTTON_BG, relief="raised", bd=1,
                          padx=8, pady=2)
            b.pack(side="right", padx=4, pady=5)
            self.toolbar_buttons.append(b)

    def _build_board(self):
        wrap = tk.Frame(self, bg=BOARD_BG)
        wrap.pack(side="top", fill="both", expand=True)
        self.board_wrap = wrap
        self.canvas = tk.Canvas(wrap, width=BOARD_W, height=BOARD_H,
                                bg=BOARD_BG, highlightthickness=0)
        self.canvas.pack(expand=True)
        # clicking a player's box picks the leader of the next trick
        self.canvas.bind("<Button-1>", self.on_canvas_click)
        # double-clicking the board brings a closed helper window back
        self.canvas.bind("<Double-Button-1>", lambda e: self.sync_popup())

    def _build_status(self):
        bar = tk.Frame(self, bg=STATUS_BG)
        bar.pack(side="top", fill="x")
        self.status_bar = bar
        self.status_var = tk.StringVar(value="")
        self.status2_var = tk.StringVar(value="")
        tk.Label(bar, textvariable=self.status_var, bg=STATUS_BG, anchor="w",
                 font=F_SMALL).pack(side="left", padx=8, pady=3)
        tk.Label(bar, textvariable=self.status2_var, bg=STATUS_BG, anchor="e",
                 font=F_SMALL, fg="#444444").pack(side="right", padx=8, pady=3)

    def _center_window(self):
        self.update_idletasks()
        w = max(self.winfo_reqwidth(), self.canvas.winfo_reqwidth())
        h = self.winfo_reqheight()
        x = max(0, (self.winfo_screenwidth() - w) // 2)
        y = max(0, (self.winfo_screenheight() - h) // 6)
        self.geometry("%dx%d+%d+%d" % (w, h, x, y))

    # -------------------------------------------------------------- helpers
    def active_seat(self):
        """Whose turn it is right now (or None)."""
        st = self.state
        if st is None:
            return None
        if st["phase"] == "bidding" and st["auction"] is not None:
            return st["auction"].turn
        if st["phase"] == "play" and st["play"] is not None:
            return st["play"].turn
        return st["dealer"] if st["dealer"] else None

    def push_undo(self):
        self.snapshots.append(copy.deepcopy(self.state))
        if len(self.snapshots) > 200:
            self.snapshots.pop(0)

    def estimates(self):
        return self.state.setdefault("estimates", dict((s, 0) for s in SEATS))

    def honor_points(self, seat):
        """The points that player has actually played (the "played" box)."""
        p = self.state["play"]
        return p.honors(seat) if p else 0

    def estimated_total(self, seat):
        """Estimate + honours -- only used for the end-of-hand summary."""
        return self.estimates().get(seat, 0) + self.honor_points(seat)

    # ---------------------------------------------------------- hand cycle
    def new_hand(self):
        prev = self.state
        hand_no = (prev["hand_no"] + 1) if prev else 1
        dealer = None
        if prev and prev["dealer"]:
            dealer = step(prev["dealer"])     # dealer rotates clockwise
        self.state = {
            "phase": "dealer",
            "hand_no": hand_no,
            "dealer": dealer,
            "next_dealer": dealer or "N",
            "auction": None,
            "play": None,
            "contract": None,
            "estimates": dict((s, 0) for s in SEATS),
        }
        self.snapshots = []
        self.sync_ui()

    def choose_dealer(self, seat):
        self.state["dealer"] = seat
        self.state["auction"] = Auction(seat)
        self.state["phase"] = "bidding"
        self.sync_ui()

    def place_bid(self, call):
        auc = self.state["auction"]
        if not auc.is_legal(call):
            return
        self.push_undo()
        auc.place(call)
        if auc.passed_out():
            self.sync_ui()
            messagebox.showinfo("Passed out",
                                "Four passes with no bid \u2014 the hand is "
                                "passed out.\n\nStart a new hand to continue.",
                                parent=self)
            self.new_hand()
            return
        self.sync_ui()

    def bidding_over(self):
        auc = self.state["auction"]
        if auc.contract() is None:
            messagebox.showinfo(
                "Nothing to play",
                "There is no contract yet \u2014 no one has made a bid.\n\n"
                "Keep bidding, or start a new hand.", parent=self)
            return
        self.push_undo()
        self.state["contract"] = auc.contract()
        self.state["phase"] = "estimates"
        self.sync_ui()

    def apply_estimates(self, values):
        self.push_undo()
        for seat in SEATS:
            self.estimates()[seat] = int(values.get(seat, 0))
        self.start_play()

    def start_play(self):
        st = self.state
        if st["contract"] is None:
            return
        leader = left_of(st["contract"]["declarer"])
        st["play"] = Play(leader, trump_of(st["contract"]))
        st["phase"] = "play"
        self.sync_ui()

    def place_card(self, suit, rank):
        p = self.state["play"]
        if p is None or p.complete():
            return
        self.push_undo()
        p.place(suit, rank)
        if p.complete():
            self.state["phase"] = "finished"
            self.sync_ui()
            self.show_summary()
            return
        self.sync_ui()

    # ------------------------------------------------------- who leads next
    def next_leader_note(self):
        """v1.7: a sentence naming the player the lead has just gone to."""
        p = self.state["play"]
        if p is None:
            return ""
        rec = p.last_record()
        if rec is None:
            return "Opening lead: %s, left of the declarer." % SEAT_NAME[p.turn]
        if not p.at_trick_start():
            return ""
        return ("Trick %d won by %s with %s \u2014 %s leads trick %d."
                % (rec["n"], SEAT_NAME[rec["winner"]], win_reason(rec),
                   SEAT_NAME[p.turn], p.trick_no()))

    def on_canvas_click(self, event):
        """Clicking a player's box starts the hand off (it picks the dealer).
        v1.7: during the play the box no longer has anything to do -- the
        lead belongs to the winner of the last trick -- so a click there
        just says so."""
        seat = seat_at(event.x, event.y)
        if seat is None:
            return
        st = self.state
        if st["phase"] == "play":
            p = st["play"]
            if p is None or p.complete():
                return
            if seat == p.turn and p.at_trick_start():
                self.status_var.set(
                    "%s leads trick %d \u2014 the winner of the last trick.  "
                    "No click is needed: enter the card."
                    % (SEAT_NAME[seat], p.trick_no()))
            else:
                self.status_var.set(
                    "%s is not on turn.  The winner of a trick leads the next "
                    "one by itself \u2014 the flashing note marks the player "
                    "whose action it is (%s to play)."
                    % (SEAT_NAME[seat], SEAT_NAME[p.turn]))
        elif st["phase"] == "dealer":
            self.push_undo()
            self.choose_dealer(seat)

    def undo(self):
        if not self.snapshots:
            return
        self.state = self.snapshots.pop()
        self.sync_ui()

    def clear_entries(self):
        if not messagebox.askyesno(
                "Clear Entries",
                "Clear the bidding, the cards played and the points\n"
                "(the date stays)?", parent=self):
            return
        st = self.state
        dealer = st["dealer"]
        hand_no = st["hand_no"]
        self.snapshots = []
        self.state = {
            "phase": "bidding" if dealer else "dealer",
            "hand_no": hand_no,
            "dealer": dealer,
            "next_dealer": dealer,
            "auction": Auction(dealer) if dealer else None,
            "play": None,
            "contract": None,
            "estimates": dict((s, 0) for s in SEATS),
        }
        self.sync_ui()

    def show_summary(self):
        st = self.state
        p = st["play"]
        lines = ["Contract: %s" % contract_text(st["contract"]),
                 "Trump: %s" % (SUIT_WORD[p.trump] if p.trump else "none (no trump)"),
                 "Opening lead: %s" % SEAT_NAME[p.leader],
                 "Tricks: N-S %d, E-W %d" % (p.side_tricks()["N-S"],
                                             p.side_tricks()["E-W"]),
                 "",
                 "%-5s  %-9s %-9s %-9s" % ("", "estimate", "played", "total")]
        for seat in SEATS:
            lines.append("%-5s  %-9d %-9d %-9d  (%d tricks)"
                         % (SEAT_NAME[seat], self.estimates()[seat],
                            p.honors(seat), self.estimated_total(seat),
                            p.tricks_won(seat)))
        messagebox.showinfo("Hand complete", "\n".join(lines), parent=self)

    # -------------------------------------------------------------- drawing
    def sync_ui(self):
        self.redraw()
        self.update_status()
        self.sync_popup()

    # ------------------------------------------------------------- flashing
    # v1.7: the note that says "on turn" (or "leads trick N") blinks, so that
    # the player whose action it is stands out at a glance.
    def start_blinking(self):
        """Start the flash if it is not already running."""
        if self.blink_job is None:
            self.blink_job = self.after(BLINK_MS, self.blink_tick)

    def blink_tick(self):
        # called by the timer -- and, in the checks, by hand: cancel whatever
        # is pending so that the flash can never end up with two timers
        if self.blink_job is not None:
            try:
                self.after_cancel(self.blink_job)
            except (tk.TclError, ValueError):
                pass
            self.blink_job = None
        self.blink_on = not self.blink_on
        self.paint_blink()
        self.start_blinking()

    def paint_blink(self):
        """Show (or hide) every flashing note, as the state is right now."""
        for item, hidden_ink in self.blink_items:
            try:
                self.canvas.itemconfigure(
                    item, fill=FLASH_TEXT if self.blink_on else hidden_ink)
            except tk.TclError:              # that item has been redrawn since
                pass

    def redraw(self):
        c = self.canvas
        c.delete("all")
        st = self.state
        active = self.active_seat()
        p = st["play"]
        self.blink_items = []            # v1.7: the old items have gone
        self.blink_on = True             # a fresh note starts visible
        for seat in SEATS:
            self.draw_panel(seat, active == seat, self.seat_note(seat, p))
        self.draw_bidding()
        self.paint_blink()

    def seat_note(self, seat, p):
        """The small note in a player's name row.  v1.7: the two that say
        whose action it is ("leads trick N" and "on turn") flash -- see
        paint_blink.  The other notes ("played 7\u2660", "won trick 2") do
        not: they are history, not a call to action."""
        st = self.state
        if p is None or st["phase"] not in ("play", "finished"):
            return ""
        if st["phase"] == "play" and not p.complete():
            if p.card_in_trick() == 1:
                if seat == p.turn:
                    return "leads trick %d" % p.trick_no()
            else:
                if seat == p.turn:
                    return "on turn"
                if seat == p.sequence[-1][0]:
                    return "played %s" % card_name(*p.sequence[-1][1:])
        rec = p.last_record()
        if rec is not None and rec["winner"] == seat:
            return "won trick %d" % rec["n"]
        return ""

    def draw_panel(self, seat, is_active, note=""):
        c = self.canvas
        x, y, w, h = PANEL_RECT[seat]
        c.create_rectangle(x, y, x + w, y + h, fill=PANEL_FILL[seat],
                           outline=ACTIVE_EDGE if is_active else PANEL_EDGE,
                           width=3 if is_active else 1)
        # name row
        c.create_text(x + PANEL_PAD + 2, y + PANEL_PAD + NAME_H // 2,
                      text=SEAT_NAME[seat], anchor="w", font=F_NAME,
                      fill="#1F3864")
        # v1.6: the top centre of every player's box says what the grid is
        c.create_text(x + w // 2, y + PANEL_PAD + NAME_H // 2,
                      text="Cards Played", font=F_SMALL, fill="#595959")
        if note:
            # v1.7: "leads trick N" and "on turn" flash (bold, and they blink
            # between FLASH_TEXT and the panel's own colour); the history
            # notes stay small and steady.
            flashes = note_flashes(note)
            item = c.create_text(x + w - PANEL_PAD - 2,
                                 y + PANEL_PAD + NAME_H // 2,
                                 text="\u25b6 " + note, anchor="e",
                                 font=F_NOTE if flashes else F_TINY,
                                 fill=NOTE_TEXT)
            if flashes:
                self.blink_items.append((item, PANEL_FILL[seat]))
        # suit rows
        top = y + PANEL_PAD + NAME_H
        for i, suit in enumerate(SUITS):
            ry = top + i * CELL_H
            c.create_text(x + PANEL_PAD + SUIT_LABEL_W - 6, ry + CELL_H // 2,
                          text="%s%s" % (SUIT_LETTER[suit], SUIT_SYMBOL[suit]),
                          anchor="e", font=F_SUIT, fill=SUIT_COLOR[suit])
            cards = self.state["play"].cards[seat][suit] if self.state["play"] else []
            extra = max(0, len(cards) - SLOTS)
            for j in range(SLOTS):
                cx = x + PANEL_PAD + SUIT_LABEL_W + j * SLOT_W
                rank = cards[j] if j < len(cards) else ""
                if extra and j == SLOTS - 1:
                    rank = "+%d" % extra          # more than ten of one suit
                honor = rank in HONOR_POINTS
                c.create_rectangle(cx, ry, cx + SLOT_W, ry + CELL_H,
                                   fill=HONOR_FILL if honor else CELL_FILL,
                                   outline=CELL_EDGE)
                if rank:
                    # v1.7: the rank is printed in the colour of its suit --
                    # red for hearts and diamonds, black for clubs and spades
                    c.create_text(cx + SLOT_W // 2, ry + CELL_H // 2, text=rank,
                                  font=F_CARD, fill=CARD_INK[suit])
                elif j == len(cards) and is_active and self.state["phase"] == "play":
                    c.create_text(cx + SLOT_W // 2, ry + CELL_H // 2, text="\u25ab",
                                  font=F_TINY, fill="#BFBFBF")
        # the two point boxes: estimated (set at the end of the bidding) and
        # played (the honours this player has put on the table)
        p = self.state["play"]
        won = p.tricks_won(seat) if p else 0
        ey = top + 4 * CELL_H + px(4)
        box_w = (PANEL_W - 2 * PANEL_PAD - POINTS_GAP) // 2
        est_x = x + PANEL_PAD
        play_x = est_x + box_w + POINTS_GAP
        c.create_text(est_x + 2, ey + EST_LABEL_H // 2, text="estimated points",
                      anchor="w", font=F_SMALL, fill="#404040")
        c.create_text(play_x + 2, ey + EST_LABEL_H // 2, text="points played",
                      anchor="w", font=F_SMALL, fill="#404040")
        c.create_text(play_x + box_w - 2, ey + EST_LABEL_H // 2,
                      text="%d trick%s" % (won, "" if won == 1 else "s"),
                      anchor="e", font=F_TINY, fill="#808080")
        for bx, value, color in ((est_x, self.estimates().get(seat, 0), "#1F3864"),
                                 (play_x, self.honor_points(seat),
                                  HONOR_TEXT if self.honor_points(seat) else "#1F3864")):
            c.create_rectangle(bx, ey + EST_LABEL_H,
                               bx + box_w, ey + EST_LABEL_H + EST_VALUE_H,
                               fill="#FFFFFF", outline="#7F7F7F", width=2)
            c.create_text(bx + box_w // 2, ey + EST_LABEL_H + EST_VALUE_H // 2,
                          text=str(value), font=F_BIG, fill=color)

    def draw_bidding(self):
        c = self.canvas
        bx, by, bw, bh = BID_RECT
        c.create_rectangle(bx, by, bx + bw, by + bh, fill="#FFFFFF",
                           outline=BID_GRID_EDGE, width=2)
        c.create_text(bx + bw // 2, by + px(8),
                      text="BIDDING",
                      font=F_SMALL, fill="#1F3864")
        auc = self.state["auction"]

        # Which seat owns which column.  If the auction has not started yet the
        # columns are laid out from the dealer, clockwise.
        if auc is not None:
            order = auc.seat_order()
        else:
            order = Auction(self.state["next_dealer"] or "N").seat_order()

        col_w = bw / 4.0
        hy = by + BID_HEADER_H
        hdr_top = hy - px(18)                 # the seat-name band, below the title
        for i, seat in enumerate(order):
            cx = bx + i * col_w
            active_col = (auc is not None and self.state["phase"] == "bidding"
                          and auc.turn == seat)
            c.create_rectangle(cx, hdr_top, cx + col_w, hy,
                               fill="#FFE699" if active_col else "#D9E1F2",
                               outline=BID_GRID_EDGE)
            c.create_text(cx + col_w / 2, (hdr_top + hy) / 2,
                          text=SEAT_NAME[seat], font=F_HEAD,
                          fill="#1F3864")

        # calls -> (round, column)
        cells = {}
        if auc is not None:
            for i, (seat, call) in enumerate(auc.calls):
                cells[(i // 4, order.index(seat))] = call
        n_rounds = max(BID_ROWS_MIN,
                       (max([r for r, _ in cells]) + 1) if cells else 0)
        avail = bh - BID_HEADER_H - 16 - 46
        pitch = min(BID_PITCH, max(px(11), avail // n_rounds))
        for r in range(n_rounds):
            ry = hy + r * pitch
            if ry + pitch > by + bh - 46:
                break
            for i in range(4):
                cx = bx + i * col_w
                c.create_rectangle(cx, ry, cx + col_w, ry + pitch,
                                   fill="#FFFFFF", outline="#D6DCE4")
                call = cells.get((r, i))
                if call:
                    c.create_text(cx + col_w / 2, ry + pitch / 2,
                                  text=bid_text(call), font=F_BID,
                                  fill="#843C0C" if call in ("X", "XX") else "#000000")
        # mark the cell the next call will go into
        if self.state["phase"] == "bidding" and auc is not None:
            r, col = len(auc.calls) // 4, order.index(auc.turn)
            ry = hy + r * pitch
            c.create_rectangle(bx + col * col_w + 1, ry + 1,
                               bx + (col + 1) * col_w - 1, ry + pitch - 1,
                               outline=ACTIVE_EDGE, width=2)

        # contract summary + trick counter under the grid
        st = self.state
        if st["contract"] is not None:
            c.create_text(bx + bw // 2, by + bh - px(26),
                          text="Contract: %s" % contract_text(st["contract"]),
                          font=F_BOLD, fill="#1F3864")
        if st["play"] is not None:
            p = st["play"]
            tricks = p.side_tricks()
            bits = ["Trump: %s" % (SUIT_SYMBOL[p.trump] if p.trump else "none")]
            bits.append("N-S %d tricks  \u00b7  E-W %d tricks"
                        % (tricks["N-S"], tricks["E-W"]))
            rec = p.last_record()
            if rec is not None:
                bits.append("trick %d won by %s" % (rec["n"], rec["winner"]))
            c.create_text(bx + bw // 2, by + bh - px(9),
                          text="   ".join(bits), font=F_SMALL, fill="#404040")

    def update_status(self):
        st = self.state
        phase = st["phase"]
        if phase == "dealer":
            self.status_var.set("Start of a hand \u2014 click the dealer's name in "
                                "the little window (or click a player's box).")
            self.status2_var.set("Hand %d" % st["hand_no"])
        elif phase == "bidding":
            auc = st["auction"]
            self.status_var.set(
                "Bidding clockwise \u2014 %s to call.   Auction so far: %s"
                % (SEAT_NAME[auc.turn], auc.auction_text() or "(nothing yet)"))
            self.status2_var.set("Hand %d   dealer %s"
                                 % (st["hand_no"], SEAT_NAME[auc.dealer]))
        elif phase == "estimates":
            self.status_var.set("The bidding is over: %s.   Enter your estimate "
                                "of each player's points, then play begins."
                                % contract_text(st["contract"]))
            self.status2_var.set("Hand %d" % st["hand_no"])
        elif phase == "play":
            p = st["play"]
            self.status2_var.set("Hand %d   opening lead %s   trump %s"
                                 % (st["hand_no"], SEAT_NAME[p.leader],
                                    SUIT_SYMBOL[p.trump] if p.trump else "none"))
            if p.card_in_trick() == 1:
                # v1.7: the lead of a trick belongs to the winner of the one
                # before it, so there is no click to make here any more.
                self.status_var.set(
                    "%s   %s   No click needed \u2014 %s is on turn (the note "
                    "on their box flashes)."
                    % (contract_text(st["contract"]), self.next_leader_note(),
                       SEAT_NAME[p.turn]))
            else:
                self.status_var.set(
                    "Card %d of 4 in trick %d \u2014 %s to play (the led suit "
                    "is pre-selected).   %s"
                    % (p.card_in_trick(), p.trick_no(), SEAT_NAME[p.turn],
                       contract_text(st["contract"])))
        elif phase == "finished":
            p = st["play"]
            tricks = p.side_tricks()
            self.status_var.set("Hand complete \u2014 %s.  All 52 cards recorded "
                                "(N-S %d tricks, E-W %d)."
                                % (contract_text(st["contract"]),
                                   tricks["N-S"], tricks["E-W"]))
            self.status2_var.set("Hand %d" % st["hand_no"])

    # --------------------------------------------------------------- popups
    def close_popup(self):
        if self.popup is not None:
            try:
                self.popup.destroy()
            except tk.TclError:
                pass
            self.popup = None

    def sync_popup(self):
        """Put the right helper window on screen for the phase we are in.

        v1.5: a window is placed once, when it first appears.  After that the
        user moves it and the app leaves it alone.
        """
        phase = self.state["phase"]
        made = False
        if phase == "dealer":
            if not isinstance(self.popup, DealerPopup):
                self.close_popup()
                self.popup = DealerPopup(self)
                made = True
        elif phase == "bidding":
            if not isinstance(self.popup, BidPopup):
                self.close_popup()
                self.popup = BidPopup(self)
                made = True
            else:
                self.popup.refresh()
        elif phase == "play":
            if not isinstance(self.popup, CardPopup):
                self.close_popup()
                self.popup = CardPopup(self)
                made = True
            else:
                self.popup.refresh()
        elif phase == "estimates":
            if not isinstance(self.popup, EstimatesPopup):
                self.close_popup()
                self.popup = EstimatesPopup(self)
                made = True
        else:
            self.close_popup()
        if made and self.popup is not None:
            self.popup.place_at_start()

    def show_help(self):
        """A small help sheet; it opens at the upper right like every other
        helper window and can be dragged out of the way."""
        if getattr(self, "help_win", None) is not None:
            try:
                if self.help_win.winfo_exists():
                    self.help_win.lift()
                    return
            except tk.TclError:
                pass
        occupied = getattr(self.popup, "placed", False) is True
        self.help_win = HelpWindow(self)
        self.help_win.place_at_start(cascade=HelpWindow.CASCADE if occupied else 0)

    def show_estimates(self):
        st = self.state
        if st["play"] is None:
            messagebox.showinfo(
                "Not yet",
                "The estimated points are entered once the bidding is over.",
                parent=self)
            return
        self.close_popup()
        self.popup = EstimatesPopup(self)
        self.popup.place_at_start()

    def run_demo(self):
        """Fill in a scripted hand so the board is not empty."""
        self.choose_dealer("E")
        for call in ("1H", "1S", "2H", "P", "3H", "P", "P", "P"):
            self.place_bid(call)
            self.update()
        self.state["contract"] = self.state["auction"].contract()
        self.state["phase"] = "estimates"
        self.apply_estimates({"N": 10, "E": 12, "S": 8, "W": 10})
        # trick 1: spades led, no trump played -- the ace wins it for North
        # trick 2: hearts led (trump) -- the jack wins it
        for suit, rank in (("S", "K"), ("S", "3"), ("S", "A"), ("S", "7"),
                           ("H", "2"), ("H", "5"), ("H", "J"), ("H", "4")):
            # v1.7: nothing has to be clicked between tricks -- the winner of
            # a trick leads the next one by itself
            self.place_card(suit, rank)
            self.update()


# --------------------------------------------------------------------------
#  Pop-up windows -- all of them are small, all of them start at the upper
#  right of the board (and never cover it), and all of them can be dragged.
# --------------------------------------------------------------------------

class BasePopup(tk.Toplevel):
    """A small, non-resizable helper window.  It starts at the upper right of
    the board and is then left wherever the user drags it (drag the title
    bar as usual)."""

    PAD = 8
    TITLE_BAR = 34          # room left for the window's own title bar
    OFFSET = 6
    CASCADE = 28            # offset used when two windows share the corner

    def __init__(self, app, corner_seat=None):
        tk.Toplevel.__init__(self, app)
        self.app = app
        self.transient(app)
        self.resizable(False, False)
        self.configure(bg="#F2F2F2")
        self.corner_seat = corner_seat or POPUP_START_CORNER
        self.corner_used = None
        self.placed = False
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.bind("<Escape>", lambda e: self.on_close())

    def on_close(self):
        """Closing a helper window just hides it; the hand is not lost."""
        self.app.close_popup()
        self.app.status_var.set(
            "Helper window closed \u2014 use New Hand, Undo or Set Estimates "
            "(or double-click the board) to bring it back."
            if self.app.state["phase"] in ("bidding", "play")
            else "Ready.")

    def place_at_start(self, cascade=0):
        """Park the window at the upper right of the board.

        v1.5: this runs once, when the window first appears.  The board is
        redrawn after every click, and a window the user has dragged must
        stay where it was put -- so nothing may re-place it later.
        """
        self.update_idletasks()
        pw = self.winfo_reqwidth() + px(12)
        ph = self.winfo_reqheight() + self.TITLE_BAR
        x, y, w, h = CORNER_RECT[self.corner_seat]
        fits = pw <= w - 2 * self.OFFSET and ph <= h - 2 * self.OFFSET
        self.corner_used = self.corner_seat if fits else None
        sx = self.app.canvas.winfo_rootx() + x + self.OFFSET + cascade
        sy = self.app.canvas.winfo_rooty() + y + self.OFFSET + cascade
        # never let the start position push the window off the screen
        sx = max(0, min(sx, self.winfo_screenwidth() - pw - px(6)))
        sy = max(0, min(sy, self.winfo_screenheight() - ph - px(30)))
        self.geometry("+%d+%d" % (sx, sy))
        self.placed = True
        self.lift()

    def head(self, text):
        f = tk.Frame(self, bg=HEADER_BG)
        f.pack(fill="x")
        tk.Label(f, text=text, bg=HEADER_BG, fg="white", font=F_P_HEAD,
                 anchor="w", padx=6, pady=3).pack(fill="x")
        return f


class DealerPopup(BasePopup):
    """Who deals.  v1.5: the bidding is clockwise, so there is nothing else
    to choose here."""

    def __init__(self, app):
        BasePopup.__init__(self, app, POPUP_START_CORNER)
        self.head("New hand \u2014 who deals?")
        body = tk.Frame(self, bg="#F2F2F2")
        body.pack(padx=6, pady=4, fill="both")

        row = tk.Frame(body, bg="#F2F2F2")
        row.pack(anchor="w")
        default = app.state.get("next_dealer") or "N"
        for seat in SEATS:
            b = tk.Button(row, text=SEAT_NAME[seat], font=F_P_BTN_BIG, width=6,
                          padx=2, pady=3,
                          relief="sunken" if seat == default else "raised",
                          command=lambda s=seat: self.pick(s))
            b.pack(side="left", padx=2)
        tk.Label(body, text="The dealer calls first; the auction and the play "
                            "are always clockwise.",
                 bg="#F2F2F2", font=F_P_TINY, anchor="w",
                 justify="left", wraplength=POPUP_WRAP).pack(fill="x", pady=(4, 0))
        tk.Label(body, text="The opening lead belongs to the player on the "
                            "declarer's left.",
                 bg="#F2F2F2", font=F_P_TINY, anchor="w",
                 justify="left", wraplength=POPUP_WRAP).pack(fill="x", pady=(2, 0))

    def pick(self, seat):
        self.app.choose_dealer(seat)


class BidPopup(BasePopup):
    """Level + strain + the special calls.  Every choice is one or two clicks."""

    def __init__(self, app):
        BasePopup.__init__(self, app, POPUP_START_CORNER)
        hf = self.head("Bid")
        self.head_label = hf.winfo_children()[0]

        body = tk.Frame(self, bg="#F2F2F2")
        body.pack(padx=6, pady=4)
        self.body = body

        lvl = tk.Frame(body, bg="#F2F2F2")
        lvl.pack(anchor="w")
        tk.Label(lvl, text="Level", width=5, anchor="w", bg="#F2F2F2",
                 font=F_P_BTN_HEAD).pack(side="left")
        self.level_buttons = {}
        for lvl_n in range(1, 8):
            b = tk.Button(lvl, text=str(lvl_n), font=F_P_BTN_BIG, width=2, pady=1,
                          bg="#FFFFFF",
                          command=lambda n=lvl_n: self.click_level(n))
            b.pack(side="left", padx=1)
            self.level_buttons[lvl_n] = b

        st = tk.Frame(body, bg="#F2F2F2")
        st.pack(anchor="w")
        tk.Label(st, text="Strain", width=5, anchor="w", bg="#F2F2F2",
                 font=F_P_BTN_HEAD).pack(side="left")
        self.strain_buttons = {}
        for strain in STRAINS:
            b = tk.Button(st, text=STRAIN_SYMBOL[strain], font=F_P_BTN_BIG, width=3,
                          pady=1, bg="#FFFFFF",
                          fg=SUIT_COLOR.get(strain, "#1F3864"),
                          command=lambda s=strain: self.click_strain(s))
            b.pack(side="left", padx=1)
            self.strain_buttons[strain] = b

        special = tk.Frame(body, bg="#F2F2F2")
        special.pack(anchor="w")
        tk.Label(special, text="Call", width=5, anchor="w", bg="#F2F2F2",
                 font=F_P_BTN_HEAD).pack(side="left")
        self.pass_btn = tk.Button(special, text="Pass", font=F_P_BTN, width=5,
                                  pady=1, bg="#FFFFFF",
                                  command=lambda: self.app.place_bid("P"))
        self.pass_btn.pack(side="left", padx=1)
        self.x_btn = tk.Button(special, text="Double", font=F_P_BTN, width=6,
                               pady=1, bg="#FFFFFF",
                               command=lambda: self.app.place_bid("X"))
        self.x_btn.pack(side="left", padx=1)
        self.xx_btn = tk.Button(special, text="Redouble", font=F_P_BTN, width=8,
                                pady=1, bg="#FFFFFF",
                                command=lambda: self.app.place_bid("XX"))
        self.xx_btn.pack(side="left", padx=1)

        tk.Button(body, text="Bidding is over", font=F_P_BTN_BIG, bg="#FCE4D6",
                  pady=2, command=self.app.bidding_over).pack(fill="x", pady=(4, 0))

        self.hint = tk.Label(body, text="", bg="#F2F2F2", font=F_P_TINY,
                             justify="left", anchor="w", wraplength=POPUP_WRAP)
        self.hint.pack(fill="x", pady=(3, 0))

        self.pending_level = None
        self.refresh()

    # one click for level, one for strain -> the bid is entered
    def click_level(self, n):
        self.pending_level = n
        self.refresh()

    def click_strain(self, strain):
        lvl = self.pending_level or self.first_legal_level(strain)
        call = "%d%s" % (lvl, strain)
        if not self.app.state["auction"].is_legal(call):
            self.hint.config(text="%s is not legal here \u2014 %s"
                                  % (bid_text(call),
                                     self.app.state["auction"].illegal_reason(call)),
                             fg="#B00020")
            self.pending_level = None
            self.refresh()
            return
        self.pending_level = None
        self.app.place_bid(call)

    def first_legal_level(self, strain):
        auc = self.app.state["auction"]
        for n in range(1, 8):
            if auc.is_legal("%d%s" % (n, strain)):
                return n
        return 7

    def refresh(self):
        auc = self.app.state["auction"]
        seat = auc.turn
        nxt = step(seat)
        self.head_label.config(text="Bid \u2014 %s   (dealer %s, clockwise)"
                                    % (SEAT_NAME[seat], SEAT_NAME[auc.dealer]))
        for n, b in self.level_buttons.items():
            legal = any(auc.is_legal("%d%s" % (n, s)) for s in STRAINS)
            selected = (self.pending_level == n)
            b.config(state="normal" if (legal or selected) else "disabled",
                     bg=PICK_FILL if selected else "#FFFFFF")
        for strain, b in self.strain_buttons.items():
            lvl = self.pending_level
            if lvl is None:
                ok = any(auc.is_legal("%d%s" % (n, strain)) for n in range(1, 8))
            else:
                ok = auc.is_legal("%d%s" % (lvl, strain))
            b.config(state="normal" if ok else "disabled")
        self.pass_btn.config(state="normal")
        self.x_btn.config(state="normal" if auc.is_legal("X") else "disabled")
        self.xx_btn.config(state="normal" if auc.is_legal("XX") else "disabled")
        last = auc.last_call()
        self.hint.config(
            text=("Click a level then a strain.  Last call: %s.  Next: %s."
                  % (bid_text(last) if last else "none", SEAT_NAME[nxt])),
            fg="#404040")


class CardPopup(BasePopup):
    """Suit then rank; the card is written into the player's panel.

    v1.5: the suit led in the trick is pre-selected, so following suit is one
    click.  Another suit can be chosen for that card, and as soon as the card
    is in, the next player is back on the led suit.

    v1.7: the leader of the trick is worked out by the app -- the winner of
    the trick before it -- and no click is needed anywhere: the row of names
    only reports who leads, and the flashing note on that player's box says
    whose action it is.
    """

    def __init__(self, app):
        BasePopup.__init__(self, app, POPUP_START_CORNER)
        hf = self.head("Card played")
        self.head_label = hf.winfo_children()[0]

        body = tk.Frame(self, bg="#F2F2F2")
        body.pack(padx=6, pady=4)
        self.body = body

        # -- who leads this trick ------------------------------------------
        # v1.7: the lead follows the cards, so this row only reports it --
        # the leader's chip is highlighted and there is nothing to click.
        lf = tk.Frame(body, bg="#F2F2F2")
        lf.pack(anchor="w")
        self.lead_row = lf
        self.lead_label = tk.Label(lf, text="Leads trick", width=11, anchor="w",
                                   bg="#F2F2F2", font=F_P_BTN_HEAD)
        self.lead_label.pack(side="left")
        self.lead_chips = {}
        for seat in SEATS:
            chip = tk.Label(lf, text=SEAT_NAME[seat], font=F_P_BTN, width=6,
                            pady=1, bd=1, relief="raised", bg="#FFFFFF")
            chip.pack(side="left", padx=1)
            self.lead_chips[seat] = chip
        # one compact note line under the leader row
        self.note = tk.Label(body, text="", bg="#F2F2F2", font=F_P_TINY,
                             justify="left", anchor="w",
                             wraplength=POPUP_WRAP)
        self.note.pack(fill="x", pady=(2, 2))

        # -- the card ------------------------------------------------------
        sf = tk.Frame(body, bg="#F2F2F2")
        sf.pack(anchor="w")
        tk.Label(sf, text="Suit", width=5, anchor="w", bg="#F2F2F2",
                 font=F_P_BTN_HEAD).pack(side="left")
        self.suit_buttons = {}
        for suit in SUITS:
            b = tk.Button(sf, text="%s%s" % (SUIT_SYMBOL[suit], SUIT_LETTER[suit]),
                          font=F_P_BTN, width=3, pady=1, bg="#FFFFFF",
                          fg=SUIT_COLOR[suit], command=lambda s=suit: self.click_suit(s))
            b.pack(side="left", padx=1)
            self.suit_buttons[suit] = b

        rf = tk.Frame(body, bg="#F2F2F2")
        rf.pack(anchor="w")
        tk.Label(rf, text="Card", width=5, anchor="w", bg="#F2F2F2",
                 font=F_P_BTN_HEAD).pack(side="left")
        self.rank_buttons = {}
        for i, rank in enumerate(RANKS):
            if i == 7:
                rf = tk.Frame(body, bg="#F2F2F2")
                rf.pack(anchor="w")
                tk.Label(rf, text="", width=5, bg="#F2F2F2",
                         font=F_P_BTN_HEAD).pack(side="left")
            b = tk.Button(rf, text=rank, font=F_P_BTN_BIG, width=2, pady=1,
                          bg="#FFFFFF",
                          command=lambda r=rank: self.click_rank(r))
            b.pack(side="left", padx=1)
            self.rank_buttons[rank] = b

        foot = tk.Frame(body, bg="#F2F2F2")
        foot.pack(anchor="w", pady=(3, 0))
        tk.Label(foot, text="Fix", width=5, anchor="w", bg="#F2F2F2",
                 font=F_P_BTN_HEAD).pack(side="left")
        tk.Button(foot, text="Undo last", font=F_P_BTN, width=9, pady=1,
                  command=self.app.undo).pack(side="left", padx=1)
        tk.Button(foot, text="Finish hand", font=F_P_BTN, width=10, pady=1,
                  command=self.finish).pack(side="left", padx=1)
        # v1.7: no "Winner leads next trick" tick any more -- the winner of a
        # trick always leads the next one, so the tick had nothing to choose.
        self.hint = tk.Label(body, text="", bg="#F2F2F2", font=F_P_TINY,
                             justify="left", anchor="w", wraplength=POPUP_WRAP)
        self.hint.pack(fill="x", pady=(2, 0))

        self.pending_suit = None
        self.suit_override = False
        self.last_turn = None
        self.last_count = -1
        self.refresh()

    def finish(self):
        if messagebox.askyesno("Finish hand",
                               "Stop recording this hand?", parent=self):
            self.app.state["phase"] = "finished"
            self.app.sync_ui()
            self.app.show_summary()

    # -- suit ---------------------------------------------------------------
    def default_suit(self):
        """The suit led in the trick being played (v1.5)."""
        p = self.app.state["play"]
        return p.led_suit()

    def click_suit(self, suit):
        p = self.app.state["play"]
        if suit == self.default_suit():
            self.suit_override = False        # back to the led suit
        else:
            self.suit_override = True         # a different suit, this card only
        self.pending_suit = suit
        self.refresh()

    def click_rank(self, rank):
        """Record one card.  v1.7: the lead is automatic, so a card is always
        accepted from the player whose turn it is; the suit must still be
        picked (it is already pre-selected for everybody who follows the led
        suit), and a card that any player has already played cannot be
        recorded at all -- the app never puts a card into a suit the user did
        not click."""
        p = self.app.state["play"]
        suit = self.pending_suit
        if suit is None:
            self.refresh()
            self.hint.config(text="Pick the suit first \u2014 then the card.",
                             fg="#B00020")
            return
        if (suit, rank) in p.played():
            self.refresh()
            self.hint.config(text="%s has already been played by somebody \u2014 "
                                  "pick another card." % card_name(suit, rank),
                             fg="#B00020")
            return
        self.app.place_card(suit, rank)

    # -- redraw -------------------------------------------------------------
    def refresh(self):
        st = self.app.state
        p = st["play"]
        seat = p.turn
        key = (seat, len(p.sequence))
        if key != (self.last_turn, self.last_count):
            # a different player (or a new trick): back to the default suit
            self.last_turn, self.last_count = key
            self.suit_override = False
            self.pending_suit = None
        default = self.default_suit()
        if not self.suit_override:
            self.pending_suit = default

        # v1.7: the lead of a trick belongs to the winner of the one before
        # it, so the window says what happened and who is on turn -- there is
        # no question to answer and nothing to click.
        at_start = p.at_trick_start()
        rec = p.last_record()
        if p.complete():
            head = "Hand complete \u2014 %s" % contract_text(st["contract"])
        elif at_start and rec is not None:
            head = "Trick %d \u2014 %s leads   (%s)" % (
                p.trick_no(), SEAT_NAME[p.turn], contract_text(st["contract"]))
        elif at_start:
            head = "Opening lead \u2014 %s   (%s)" % (
                SEAT_NAME[p.turn], contract_text(st["contract"]))
        else:
            head = "Card %d of 4 \u2014 %s to play   (%s)" % (
                p.card_in_trick(), SEAT_NAME[p.turn],
                contract_text(st["contract"]))
        self.head_label.config(text=head)

        # the leader row: it only reports the lead (there is nothing to click)
        self.lead_label.config(text="Leads trick" if at_start else "Led by")
        for s, chip in self.lead_chips.items():
            here = (s == p.turn)
            chip.config(bg=PICK_FILL if here else "#FFFFFF",
                        relief="sunken" if here else "raised",
                        fg="#1F3864" if here else "#808080")

        # one compact note: what just happened, and what the suit row means
        parts = []
        if p.complete():
            parts.append("The hand is complete.")
        elif at_start:
            if self.app.next_leader_note():
                parts.append(self.app.next_leader_note())
            parts.append("Nothing to click \u2014 the winner of a trick leads "
                         "the next one.")
        else:
            lead_card = p.trick_cards()[0]
            parts.append("%s led %s.  Undo takes the trick's cards back."
                         % (SEAT_NAME[lead_card[0]],
                            card_name(lead_card[1], lead_card[2])))
        if not p.complete():
            if default is None:
                parts.append("You lead \u2014 pick the suit, then the card.")
            elif self.suit_override:
                parts.append("Changed to %s for this card; the next player "
                             "returns to %s."
                             % (SUIT_WORD[self.pending_suit],
                                SUIT_SYMBOL[default]))
            else:
                parts.append("Led suit %s is pre-selected." % SUIT_SYMBOL[default])
        self.note.config(text="  ".join(parts),
                         fg="#B00020" if (self.suit_override
                                          and not p.complete())
                         else "#404040")

        # the suit row: a suit whose thirteen cards have all been played
        # cannot be chosen any more
        played = p.played()
        for suit, b in self.suit_buttons.items():
            gone = all((suit, r) in played for r in RANKS)
            if self.suit_override and suit == self.pending_suit:
                b.config(bg=PICK_FILL, relief="sunken")
            elif suit == default:
                b.config(bg=LED_FILL, relief="sunken")
            else:
                b.config(bg="#FFFFFF", relief="raised")
            b.config(state="disabled" if gone else "normal")

        # the rank row: v1.7 -- a rank is clickable when the card in the suit
        # that is selected is still in the pack.  Nothing is recorded into a
        # suit the user has not clicked, and a card somebody has already
        # played can never be clicked.
        for rank, b in self.rank_buttons.items():
            live = (not p.complete()
                    and self.pending_suit is not None
                    and (self.pending_suit, rank) not in played)
            b.config(state="normal" if live else "disabled",
                     bg=PICK_FILL if live else "#FFFFFF")
        if self.pending_suit is None:
            self.hint.config(
                text="Pick the suit, then the card  \u00b7  A/K/Q/J = 4/3/2/1 "
                     "points", fg="#B00020")
        else:
            self.hint.config(
                text=("Trick %d, card %d of 4  \u00b7  entering %s cards  "
                      "\u00b7  A/K/Q/J = 4/3/2/1 points"
                      % (p.trick_no(), p.card_in_trick(),
                         SUIT_SYMBOL[self.pending_suit])),
                fg="#404040")


class EstimatesPopup(BasePopup):
    """The user's estimate of each player's points, entered at the end of the
    bidding.  v1.5: the estimate box is set once and stays as it is once play
    starts; the honours a player actually plays are kept separately in the
    "points played" box."""

    def __init__(self, app):
        BasePopup.__init__(self, app, POPUP_START_CORNER)
        self.head("Estimated points \u2014 end of the bidding")
        body = tk.Frame(self, bg="#F2F2F2")
        body.pack(padx=6, pady=4)
        self.vars = {}
        for seat in SEATS:
            row = tk.Frame(body, bg="#F2F2F2")
            row.pack(fill="x", pady=0)
            tk.Label(row, text=SEAT_NAME[seat], width=6, anchor="w", bg="#F2F2F2",
                     font=F_P_SMALL).pack(side="left")
            var = tk.StringVar(value=str(app.estimates().get(seat, 0)))
            self.vars[seat] = var
            tk.Button(row, text="\u2212", font=F_P_BTN, width=2, pady=1,
                      command=lambda s=seat: self.bump(s, -1)).pack(side="left")
            tk.Entry(row, textvariable=var, width=4, font=F_P_SMALL,
                     justify="center").pack(side="left", padx=2)
            tk.Button(row, text="+", font=F_P_BTN, width=2, pady=1,
                      command=lambda s=seat: self.bump(s, +1)).pack(side="left")
        tk.Label(body, text="Fixed once play begins \u2014 the honours a player "
                            "plays go in their own box.",
                 bg="#F2F2F2", font=F_P_TINY, anchor="w", justify="left",
                 wraplength=POPUP_WRAP).pack(fill="x", pady=(3, 0))
        row = tk.Frame(body, bg="#F2F2F2")
        row.pack(fill="x", pady=(3, 0))
        if app.state["play"] is not None:
            tk.Button(row, text="Apply", font=F_P_BTN_BIG, width=10, pady=1,
                      command=self.apply).pack(side="left", padx=2)
            tk.Button(row, text="Cancel", font=F_P_BTN, width=10, pady=1,
                      command=self.cancel).pack(side="left", padx=2)
        else:
            tk.Button(row, text="Start play", font=F_P_BTN_BIG, width=12, pady=1,
                      command=self.apply).pack(side="left", padx=2)
            tk.Button(row, text="Skip (all 0)", font=F_P_BTN, width=12, pady=1,
                      command=self.skip).pack(side="left", padx=2)

    def bump(self, seat, delta):
        try:
            v = int(self.vars[seat].get())
        except ValueError:
            v = 0
        self.vars[seat].set(str(max(0, min(40, v + delta))))

    def values(self):
        out = {}
        for seat, var in self.vars.items():
            try:
                out[seat] = max(0, min(40, int(var.get())))
            except ValueError:
                out[seat] = 0
        return out

    def apply(self):
        st = self.app.state
        if st["play"] is None:
            self.app.apply_estimates(self.values())
        else:
            self.app.push_undo()
            for seat, v in self.values().items():
                st["estimates"][seat] = v
            self.app.close_popup()
            self.app.sync_ui()

    def skip(self):
        self.app.apply_estimates(dict((s, 0) for s in SEATS))

    def cancel(self):
        """Close without changing anything and bring the hand's own window
        back."""
        self.app.close_popup()
        self.app.sync_ui()


class HelpWindow(BasePopup):

    TEXT = """BRIDGE TRACKER v1.7 - quick guide

START OF A HAND
  A small window asks who deals.  Click North, East, South or West (or
  click the player's box on the board).  Bidding and play always go
  clockwise - there is no direction to choose.  The board's middle block
  is the auction: one column per player, one row per round.

BIDDING
  Click a level (1-7), then a strain (C D H S NT).  Pass / Double /
  Redouble are single clicks.  Calls that are not legal are greyed out.
  "Bidding is over" closes the auction.

END OF THE BIDDING
  Enter your estimate of each player's points with the + and - buttons (or
  by typing).  The estimate appears in that player's "estimated points"
  box and is fixed from then on - the honours that get played go into the
  separate "points played" box.

PLAY
  The opening lead belongs to the player on the declarer's left; play then
  runs clockwise, one player at a time.  The suit that was led is already
  selected in the card window, so following suit is a single click on the
  card.  Clicking a different suit changes it for that one card, and the
  next player starts again on the led suit.  A card that somebody has
  already played is greyed out and cannot be clicked, and the card window
  never puts a card into a suit you have not clicked.

WHO LEADS THE NEXT TRICK
  A trick is four cards.  When the fourth card is in, the app works the
  winner out - highest trump if the contract has a trump suit, otherwise
  the highest card of the suit that was led - says who won it and with
  which card, and gives the lead of the next trick straight to that
  player.  There is nothing to click: the winner of a trick always leads
  the next one, and the card window's row of names just reports who that
  is.

WHO IS ON TURN
  The player whose card is wanted is ringed in orange, and the note in the
  corner of their box - "leads trick N" or "on turn" - FLASHES.  The cards
  a player has played are printed in the colour of their suit: red for
  hearts and diamonds, black for clubs and spades (an honour keeps its
  tinted cell).

POINTS
  An ace adds 4, a king 3, a queen 2 and a jack 1 point to that player's
  "points played" box; it happens the moment you click the card.  The
  small note in the corner of that box shows the tricks that player won.

BUTTONS
  New Hand          ask for a new dealer and clear the board
  Undo              step back one call or one card (repeatable)
  Clear Entries     wipe the bidding, cards and points (like the macro)
  Set Estimates     re-open the estimate window
  Help              this window

Every helper window can be dragged by its title bar.  They always open at
the upper right of the board, and closing one never loses the hand.
"""

    def __init__(self, app):
        BasePopup.__init__(self, app, POPUP_START_CORNER)
        self.head("Help \u2014 Bridge Tracker v1.7")
        body = tk.Frame(self, bg="#F2F2F2")
        body.pack(padx=6, pady=4)
        txt = tk.Text(body, width=46, height=10, font=F_P_TINY, wrap="word",
                      bg="#FFFFFF", relief="solid", bd=1)
        txt.pack(side="left")
        sb = tk.Scrollbar(body, command=txt.yview)
        sb.pack(side="right", fill="y")
        txt.config(yscrollcommand=sb.set)
        txt.insert("1.0", self.TEXT)
        txt.config(state="disabled")
        tk.Button(body, text="Close", font=F_P_BTN, width=8, pady=1,
                  command=self.close_help).pack(side="bottom", pady=(4, 0))

    def close_help(self):
        self.app.help_win = None
        self.destroy()


# --------------------------------------------------------------------------

def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    if "--version" in argv or "-V" in argv:
        print("Bridge Tracker %s" % __version__)
        return 0
    demo = "--demo" in argv or "-d" in argv
    app = BridgeTracker(demo=demo)
    app.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
