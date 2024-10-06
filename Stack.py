import linked_list

class Stack:
    def __init__(self):
        self.linked_list = linked_list.LinkedList()

    def is_empty(self):
        return self.linked_list.is_empty()

    def push(self, data):
        self.linked_list.prepend(data)

    def pop(self):
        if self.is_empty():
            raise IndexError("pop from an empty stack")
        data = self.linked_list.head.data
        self.linked_list.delete(data)
        return data

    def peek(self):
        if self.is_empty():
            raise IndexError("peek from an empty stack")
        return self.linked_list.head.data

    def display(self):
        self.linked_list.display()

