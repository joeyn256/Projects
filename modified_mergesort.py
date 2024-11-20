# Python program for implementation of MergeSort
 
# Merges two subarrays of arr[].
# First subarray is arr[l..m]
# Second subarray is arr[m+1..r]
 
 
def merge(arr, l, m, r):
    n1 = m - l + 1 #n1 is length of temp array L, l is first index in the arry that we are merging, m is first index of second array, r is length of the total array
    n2 = r - m
 
    # create temp arrays
    L = [0] * (n1)
    R = [0] * (n2)
 
    # Copy data to temp arrays L[] and R[]
    for i in range(0, n1):
        L[i] = arr[l + i]
 
    for j in range(0, n2):
        R[j] = arr[m + 1 + j]
 
    # Merge the temp arrays back into arr[l..r]
    i = 0     # Initial index of first subarray
    j = 0     # Initial index of second subarray
    k = l     # Initial index of merged subarray
 
    while i < n1 and j < n2:
        if L[i] <= R[j]:
            arr[k] = L[i]
            i += 1
        else:
            arr[k] = R[j]
            j += 1
        k += 1
 
    # Copy the remaining elements of L[], if there
    # are any
    while i < n1:
        arr[k] = L[i]
        i += 1
        k += 1
 
    # Copy the remaining elements of R[], if there
    # are any
    while j < n2:
        arr[k] = R[j]
        j += 1
        k += 1

# This code is contributed by Mohit Kumra

def recursive_merge(lists, left, right):
    # Base case: if there's only one list, return it as-is
    if left == right:
        return lists[left]

    # Find the middle index to split lists into two halves
    mid = (left + right) // 2

    # Recursively merge left and right halves
    left_merged = recursive_merge(lists, left, mid)
    right_merged = recursive_merge(lists, mid + 1, right)

    # Merge the two halves into a single sorted array
    merged_list = left_merged + right_merged
    merge(merged_list, 0, len(left_merged) - 1, len(merged_list) - 1)

    return merged_list

# Example usage
lists = [
    [1, 5],
    [2, 3, 8, 11, 15, 21, 24],
    [50],
    [11, 12, 15, 17, 18, 19, 22],
    [3, 3, 8, 23],
    [-1,4,4,18,19]
    # Add more sorted lists if needed, up to 21 or more
]

# Start the recursive merge on the list of sorted lists
merged_list = recursive_merge(lists, 0, len(lists) - 1)
print("Merged list:", merged_list)

"""
def mergeSort(arr, l, r):
    if l < r:
 
        # Same as (l+r)//2, but avoids overflow for
        # large l and h
        m = l+(r-l)//2
 
        # Sort first and second halves
        mergeSort(arr, l, m) # if this goes to if and not true checks next
        mergeSort(arr, m+1, r) # if above wasnt true has to check this before
        merge(arr, l, m, r)

all = [a1, a2, a3, a4, a5]
all_len = len(all)
ind_lists = [0] * all_len 

ind_lists[0] = len(all[0])

for i in range(all_len - 1):
    ind_lists[i+1] = len(all[i + 1]) + ind_lists[i]
# new variable len_lists holds the lenths of all the lists
#print(ind_lists)
print(ind_lists)
m = all_len // 2 


# Driver code to test above
arr = [1, 5, 2, 3, 8, 11, 15, 21, 24]
n = len(arr)
print("Given array is")
for i in range(n):
    print("%d" % arr[i],end=" ")
 
merge(arr, 0, 1, 8)
print("\n\nSorted array is")
for i in range(n):
    print("%d" % arr[i],end=" ")

arr = [1, 5, 2, 3, 8, 11, 15, 21, 24, 15, 11, 12, 15, 17, 18, 19, 22, 3, 3, 8, 23]
#      [0  1][2  3  4  5   6   7   8]  [9][10, 11, 12, 13, 14, 15, 16]

ind_lists = [0, 2, 9, 10, 17, 21]
merge(arr, 0, 1, 8)
print(arr)

merge(arr, 9, 9, 16)
print(arr)

merge(arr, 0, 8, 16)
print(arr)

merge(arr, 0, 16, 20)
print(arr)

def iterative_merge(arr):
    # Start with the full array of sorted lists
    while len(arr) > 1:
        merged_round = []
        
        # Merge lists in pairs
        for i in range(0, len(arr), 2):
            if i + 1 < len(arr):
                # Merge current list with the next list
                merged_round.append(merge_two(arr[i], arr[i + 1]))
            else:
                # If there's an odd list out, add it directly to the next round
                merged_round.append(arr[i])
        
        # Prepare for the next round
        arr = merged_round
    return arr[0] if arr else []

def merge_two(list1, list2):
    # Merges two sorted lists and returns a single sorted list
    result = []
    i, j = 0, 0
    
    while i < len(list1) and j < len(list2):
        if list1[i] <= list2[j]:
            result.append(list1[i])
            i += 1
        else:
            result.append(list2[j])
            j += 1

    # Append remaining elements
    result.extend(list1[i:])
    result.extend(list2[j:])
    
    return result

# Example usage with 5 sorted lists
lists = [
    [1, 5],
    [2, 3, 8, 11, 15, 21, 24],
    [15],
    [11, 12, 15, 17, 18, 19, 22],
    [3, 3, 8, 23]
]



# Run iterative merge
merged_list = iterative_merge(lists)
print("Merged list:", merged_list)
        
"""