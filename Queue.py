import linked_list
import pickle

class Queue:
    def __init__(self):
        self.linked_list = linked_list.LinkedList()

    def is_empty(self):
        return self.linked_list.is_empty()

    def enqueue(self, data):
        self.linked_list.append(data)

    def dequeue(self):
        if self.is_empty():
            raise IndexError("dequeue from an empty queue")
        data = self.linked_list.head.data
        self.linked_list.head = self.linked_list.head.next_node
        return data

    def front(self):
        if self.is_empty():
            raise IndexError("front from an empty queue")
        return self.linked_list.head.data

    def display(self):
        self.linked_list.display()

    def serialize(self, file_path):
        # Serialize the entire Queue object using pickle
        with open(file_path, 'wb') as file:
            pickle.dump(self, file)

    @classmethod
    def deserialize(cls, file_path):
        # Deserialize the Queue object from the file using pickle
        with open(file_path, 'rb') as file:
            return pickle.load(file)

'''
restored.display()
# Example usage:
queue = Queue()
queue.enqueue(1)
queue.enqueue(2)
queue.enqueue(3)

print("Queue:")
queue.display()  # Output: 1 -> 2 -> 3 -> None

print("Front:", queue.front())  # Output: Front: 1

dequeued_element = queue.dequeue()
print("Dequeued element:", dequeued_element)  # Output: Dequeued element: 1

print("Queue after dequeue:")
queue.display()  # Output: 2 -> 3 -> None
'''