from Stack import Stack

class sorting_functions():

    #Sorting functions for orders
    #Returns orders based on ruppee value or status
    def number_sort(self, data, ascending):
        n = len(data)

        for j in range(n):
            swapped = False
            for i in range(0, n - j - 1):
                if data[i] > data[i + 1]:
                    data[i], data[i + 1] = data[i + 1], data[i]
                    swapped = True

            if not swapped:
                # If no swaps were made, the list is already sorted
                break

        sorted_data = data

        if not ascending:
            stack = Stack()

            # Pass data through stack class
            for k in range(0, n):
                stack.push(data[k])

            sorted_data = []

            for l in range(0, n):
                stack_data = str(stack.pop())
                sorted_data.append(stack_data)

        return sorted_data

    #Returns order based on due date
    def date_sort(self, date_strings, ascending):

        # Function to convert date string to a list of integers [day, month, year]
        def convert_to_list(date):
            return [int(date[:2]), int(date[2:4]), int(date[4:])]

        # Split and convert date strings to lists
        date_lists = [convert_to_list(date) for date in date_strings]

        # Sort by years
        n = len(date_lists)
        for i in range(n):
            for j in range(0, n - i - 1):
                if date_lists[j][2] > date_lists[j + 1][2]:
                    date_lists[j], date_lists[j + 1] = date_lists[j + 1], date_lists[j]

        # Sort by months
        for i in range(n):
            for j in range(0, n - i - 1):
                if date_lists[j][1] > date_lists[j + 1][1] and date_lists[j][2] == date_lists[j + 1][2]:
                    date_lists[j], date_lists[j + 1] = date_lists[j + 1], date_lists[j]

        # Sort by days
        for i in range(n):
            for j in range(0, n - i - 1):
                if date_lists[j][0] > date_lists[j + 1][0] and date_lists[j][1] == date_lists[j + 1][1] and date_lists[j][2] == date_lists[j + 1][2]:
                    date_lists[j], date_lists[j + 1] = date_lists[j + 1], date_lists[j]

        # Convert back to date strings
        sorted_date_strings = ['{:02d}{:02d}{:04d}'.format(day, month, year) for [day, month, year] in date_lists]

        sorted_data = sorted_date_strings

        if not ascending:
            stack = Stack()

            # Pass data through stack class
            for k in range(0, n):
                stack.push(sorted_date_strings[k])

            sorted_data = []

            for l in range(0, n):
                stack_data = str(stack.pop())
                sorted_data.append(stack_data)

        return sorted_data

    #Returns past orders of specific customer/vendor in alphabetical order
    def alphabetical_sort(self, words, ascending):

        n = len(words)

        for i in range(1, len(words)):
            key = words[i]
            j = i - 1
            while j >= 0 and key < words[j]:
                words[j + 1] = words[j]
                j -= 1
            words[j + 1] = key

        sorted_data = words

        if not ascending:
            stack = Stack()

            # Pass data through stack class
            for k in range(0, n):
                stack.push(words[k])

            sorted_data = []

            for l in range(0, n):
                stack_data = str(stack.pop())
                sorted_data.append(stack_data)

        return sorted_data