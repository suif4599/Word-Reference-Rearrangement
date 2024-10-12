import win32com.client as win32
import re
import os
from typing import List, Tuple, Dict, Set
from math import inf

class Word:
    def __init__(self, path: str | None = None, visible: bool = False):
        self.word = win32.gencache.EnsureDispatch('Word.Application')
        self.word.Visible = visible
        if path is not None:
            path = os.path.abspath(path)
            self.doc = self.word.Documents.Open(path)
        else:
            self.doc = self.word.Documents.Add()

    def save(self, path: str | None = None):
        if path is not None:
            path = os.path.abspath(path)
            self.doc.SaveAs(path)
        else:
            self.doc.Save()

    def close(self):
        self.doc.Close()

    def quit(self):
        self.word.Quit()

class WordPaper:
    "the class suppose references are in the fotmat of "
    "'[<num>]' in the text and '[<num>]: ' in the references"
    def __init__(self, path: str, visible: bool = False):
        self.word = Word(path, visible=visible)
        self.ref_in_text: List[Tuple[int, Tuple[int, int]]] = []
        self.ref_in_ref: Dict[int, Tuple[int, int]] = {}
        self.ref_start_para: int = inf
        self.__scan_references()
        self.text_sub: List[Tuple[int, int, str]] = []


    def __scan_references(self):
        for para_ind, para in enumerate(self.word.doc.Paragraphs, 1):
            rng = para.Range
            st, ed = rng.Start, rng.End
            text = rng.Text
            for match in re.finditer(r"\[\d+\]\:?", text):
                pos = st + match.start(), st + match.end()
                text = match.group()
                num = int(re.search(r"\d+",text).group())
                if text[-1] == ':':
                    self.ref_in_ref[num] = pos[0], pos[1] - 1
                    if para_ind < self.ref_start_para:
                        self.ref_start_para = para_ind
                else:
                    self.ref_in_text.append((num, pos))
    
    def __intercept(self, old_st: int, old_ed: int, new_st: int, new_ed: int) -> bool:
        "check if the two ranges are intersected"
        return old_st <= new_st <= old_ed or old_st <= new_ed <= old_ed \
            or new_st <= old_st <= new_ed or new_st <= old_ed <= new_ed

    def __register_sub_text(self, st: int, ed: int, sub: str):
        "register a text substitution action without caring the position change"
        for ost, oed, _ in self.text_sub:
            if self.__intercept(ost, oed, st, ed):
                raise ValueError("Some part of the new text has been registered")
        self.text_sub.append((st, ed, sub))

    def __swap_paragraph(self, para1: int, para2: int):
        "swap two paragraphs"
        p1 = self.word.doc.Paragraphs(para1)
        p2 = self.word.doc.Paragraphs(para2)
        p1rng, p2rng = p1.Range, p2.Range
        p1rng.Copy()
        self.word.doc.Range(p2rng.End, p2rng.End).Paste()
        p2rng.Cut()
        p1rng.Paste()


    def re_arrange_ref(self):
        "rearrange the references in the text"
        occurred: Set[int] = set()
        ind = 0
        mapping: Dict[int, int] = {} # num -> ind
        for num, (st, ed) in self.ref_in_text:
            # rng = self.word.doc.Range(st + 1, ed - 1) # exclude the brackets
            if num not in occurred:
                occurred.add(num)
                ind += 1
                mapping[num] = ind
                # rng.Text = str(ind)
                self.__register_sub_text(st + 1, ed - 1, str(ind))
                st, ed = self.ref_in_ref[num]
                # rng = self.word.doc.Range(st + 1, ed - 1)
                # rng.Text = str(ind)
                self.__register_sub_text(st + 1, ed - 1, str(ind))
            else:
                # rng.Text = str(mapping[num])
                self.__register_sub_text(st + 1, ed - 1, str(mapping[num]))
    
    def apply_sub(self):
        "apply the registered text substitution actions"
        self.text_sub.sort(key=lambda x: x[0], reverse=True)
        for st, ed, sub in self.text_sub:
            rng = self.word.doc.Range(st, ed)
            rng.Text = sub
    
    def sort_ref(self):
        "sort the references in the reference section"
        num_map: Dict[int, int] = {} # para_ind -> num
        for i in range(len(self.word.doc.Paragraphs), 0, -1):
            if self.word.doc.Paragraphs(i).Range.Text.strip():
                break
            self.word.doc.Paragraphs(i).Range.Delete()
        para_len = len(self.word.doc.Paragraphs)
        for i in range(self.ref_start_para, para_len + 1):
            text = self.word.doc.Paragraphs(i).Range.Text
            match = re.search(r"\[\d+\]\:", text)
            if not match:
                if text.strip():
                    raise ValueError("The reference section is not well formatted")
                else:
                    raise ValueError("Don't leave any empty line in the reference section")
            num = int(re.search(r"\d+", match.group()).group())
            num_map[i] = num
        self.word.doc.Paragraphs.Add() # a new empty paragraph
        for i in range(self.ref_start_para, para_len + 1):
            for j in range(para_len, i, -1):
                if num_map[j] < num_map[j - 1]:
                    self.__swap_paragraph(j - 1, j)
                    num_map[j - 1], num_map[j] = num_map[j], num_map[j - 1]

    def save(self, path: str | None = None):
        self.word.save(path)

    def close(self):
        self.word.close()
    
    def quit(self):
        self.word.quit()


import argparse

parser = argparse.ArgumentParser()
parser.add_argument("path", help="The path of the word file", type=str)
parser.add_argument("-o", "--output", help="Write the result to a new file", type=str)
args = parser.parse_args()
path = args.path
output = args.output
if ' ' in path or output and ' ' in output:
    raise ValueError("pywin32 cannot process path that contain space")
w = WordPaper(path)
w.re_arrange_ref()
w.apply_sub()
w.sort_ref()
w.save(output)
w.close()
w.quit()