from pprint import pprint
def permute(nums):
    res = []
    track = []
    def backtrack(nums, track):
        if len(track) == len(nums):
            res.append([i for i in track])
            return
        for num in nums:
            if num in track:
                continue
            track.append(num)
            backtrack(nums, track)
            track.pop()
    
    backtrack(nums, track)
    pprint(res)
    print(len(res))
    pass

if __name__ == "__main__":
    permute(list(range(3, 7)))