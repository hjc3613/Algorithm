
def helper(arr, l, h):
    index = l-1
    spliter = arr[h]
    for i in range(l, h):
        if arr[i] <= spliter:
            index += 1
            if i > index:
                arr[i], arr[index] = arr[index], arr[i]
    index += 1
    arr[index], arr[h] = arr[h], arr[index]
    return index 

def quick_sort(arr, l, h):
    if l >= h:
        return
    index = helper(arr, l, h)
    quick_sort(arr, l, index-1)
    quick_sort(arr, index+1, h)

if __name__ == "__main__":
    arr = [9,8,7,6,5,4,3,2,1,0,5,4,7,5,3,8,1,4,7,6]
    quick_sort(arr, 0, len(arr)-1)
    print(arr)
