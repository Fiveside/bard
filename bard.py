#!/usr/bin/env python3
# This is a quick and dirty script that sends keys to ffxiv for the purpose
# of playing musical instruments in game.  It reads a "sheet music" file
# and plays the bound keys as appropriate to play the music
#
# The sheet music is a series of whitespace delimited keys of the form
# 4A-1 where 4 represents the type of note (4 = quarter note, 2 = half note, etc)
# A represents the note being pressed, and -1 declares it to be one octave below
# Other examples:
# B     - quarter B note, normal octave
# 32C+2 - a 32nd C note, 2 octaves up. (highest note playable in game)
# 4.    - quarter rest
# 2.    - half rest
#
# This was written in the dead of night over the course of about an hour
# its probably buggy

import win32gui
import win32con
import win32api
import itertools
import time
import re
import sys
import logging


class NoteButton:
    def __init__(self, vk, shift=False, ctrl=False):
        self.vk = vk
        self.shift = shift
        self.ctrl = ctrl

    def play(self, instrument):
        if self.vk:
            instrument.keypress(self.vk, shift=self.shift, ctrl=self.ctrl)


BASE_NOTE_MAP = {
    'C': ord('Q'),
    'C#': ord('2'),

    'D': ord('W'),

    'Eb': ord('3'),
    'E♭': ord('3'),
    'E': ord('E'),

    'F': ord('R'),
    'F#': ord('5'),

    'G': ord('T'),
    'G#': ord('6'),

    'A': ord('Y'),

    'Bb': ord('7'),
    'B♭': ord('7'),
    'B': ord('U'),
}

NOTES = dict(itertools.chain(
    # The basic notes
    ((k, NoteButton(v)) for k, v in BASE_NOTE_MAP.items()),

    # Plus one octave
    (("{}+1".format(k), NoteButton(v, shift=True))
     for k, v in BASE_NOTE_MAP.items()),

    # Minus one octave
    (("{}-1".format(k), NoteButton(v, ctrl=True))
     for k, v in BASE_NOTE_MAP.items()),
))

NOTES['C+2'] = NoteButton(ord('I'), shift=True)
NOTES['.'] = NoteButton(None)


class Note:
    def __init__(self, duration, note_button=None):
        self.duration = duration
        self.nb = note_button

    def play(self, instrument):
        self.nb.play(instrument)
        time.sleep(self.duration)


class Instrument:
    # Yes, the game window is my instrument damnit!
    def __init__(self, hwnd):
        self.hwnd = hwnd
        self.shift_pressed = False
        self.ctrl_pressed = False

    def _send_message(self, vk, keydown=True):
        if keydown:
            message = win32con.WM_KEYDOWN
            logging.debug("KEYDOWN: {} ({})".format(vk, chr(vk)))
        else:
            message = win32con.WM_KEYUP
            logging.debug("KEYUP:   {} ({})".format(vk, chr(vk)))
        win32gui.PostMessage(self.hwnd, message, vk, 0)

    def keydown(self, vk, ctrl=False, shift=False):
        if shift != self.shift_pressed:
            # this works because keydown state is what we're switching to
            self._send_message(win32con.VK_SHIFT, keydown=shift)
            self.shift_pressed = shift
        if ctrl != self.ctrl_pressed:
            self._send_message(win32con.VK_CONTROL, keydown=ctrl)
            self.ctrl_pressed = ctrl
        self._send_message(vk)
        # win32gui.SendMessage(self.hwnd, win32con.WM_KEYDOWN, vk, 0)

    def keyup(self, vk):
        self._send_message(vk, keydown=False)

    def keypress(self, vk, *args, **kwargs):
        self.keydown(vk, *args, **kwargs)
        self.keyup(vk)

    def finish(self):
        if self.shift_pressed:
            self.keyup(win32con.VK_SHIFT)
        if self.ctrl_pressed:
            self.keyup(win32con.VK_CONTROL)


def translate_music_line(line):
    for part in line.split():
        snote_type, snote = re.search(r'(\d*)(.+)', part).groups()
        if not snote_type.strip():
            note_type = 4
        else:
            note_type = int(snote_type.strip())
        logging.debug("Translating {}: type: {}, note: {}".format(
            part, note_type, snote))
        if snote not in NOTES:
            import ipdb
            ipdb.set_trace()
        yield Note(2/note_type, NOTES[snote])


if __name__ == '__main__':
    logging.basicConfig(level=logging.DEBUG)
    with open(sys.argv[1], 'r', encoding='utf8') as fobj:
        score = fobj.readlines()
    hwnd = win32gui.FindWindow('FFXIVGAME', 'FINAL FANTASY XIV')
    if not hwnd:
        raise RuntimeError("Unable to find ffxiv window, is the game running?")
    instrument = Instrument(hwnd)
    for line in score:
        if not line.strip() or line.strip().startswith("#"):
            continue
        for note in translate_music_line(line):
            note.play(instrument)
    instrument.finish()
