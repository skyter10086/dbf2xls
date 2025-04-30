import re
from . import Style

class Block:
    def __init__(
        self,
        head: str,
        tail: str,
        style: Style,
        value: list[list],
        merge=False,
    ):
        self.head = head
        self.tail = tail
        self.style = style
        self.value = value
        self.merge = merge

        p = re.compile("[A-Z][0-9]+")
        if not (p.match(head) and p.match(tail)):
            raise Exception("Head or Tail should like 'A1' 'B3'")

        if not (
            head[0] <= tail[0]
            and int(head[1:]) <= int(tail[1:])
            and int(head[1:]) > 0
            and int(tail[1:]) > 0
        ):
            raise Exception("Head should be left and up from Tail")

        self.__cols_count = ord(tail[0]) - ord(head[0]) + 1
        self.__rows_count = int(tail[1:]) - int(head[1:]) + 1

        if merge == False:
            if not (len(value) == self.__rows_count and len(value[0]) == self.__cols_count):
                raise Exception(
                        "length of value should be rows count, length of value element should be cols count"
                )
        else:
            if not (len(value)==1 and len(value[0])==1):
                raise Exception("length of value and length of value element should be 1")
