class Node:
    def __init__(self, data):
        self.data = data
        self.next_node = None

class LinkedList:
    def __init__(self):
        self.head = None

    def is_empty(self):
        return self.head is None

    def append(self, data):
        new_node = Node(data)
        if self.head is None:
            self.head = new_node
            return
        last_node = self.head
        while last_node.next_node:
            last_node = last_node.next_node
        last_node.next_node = new_node

    def prepend(self, data):
        new_node = Node(data)
        new_node.next_node = self.head
        self.head = new_node

    def delete(self, data):
        if self.head is None:
            return

        if self.head.data == data:
            self.head = self.head.next_node
            return

        current_node = self.head
        while current_node.next_node and current_node.next_node.data != data:
            current_node = current_node.next_node

        if current_node.next_node:
            current_node.next_node = current_node.next_node.next_node

    def display(self):
        current_node = self.head
        while current_node:
            print(current_node.data, end=" -> ")
            current_node = current_node.next_node
        print("End")



'''
# Example usage:
stack = Stack()
stack.push(1)
stack.push(2)
stack.push(3)

print("Stack:")
stack.display()  # Output: 3 -> 2 -> 1 -> None

print("Peek:", stack.peek())  # Output: Peek: 3

popped_element = stack.pop()
print("Popped element:", popped_element)  # Output: Popped element: 3

print("Stack after pop:")
stack.display()  # Output: 2 -> 1 -> None

# Example usage:
linked_list = LinkedList()
linked_list.append(1)
linked_list.append(2)
linked_list.append(3)
linked_list.display()  # Output: 1 -> 2 -> 3 -> End

linked_list.prepend(0)
linked_list.display()  # Output: 0 -> 1 -> 2 -> 3 -> End

linked_list.delete(2)
linked_list.display()  # Output: 0 -> 1 -> 3 -> End
'''