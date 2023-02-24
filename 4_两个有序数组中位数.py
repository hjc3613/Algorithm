
class Solution:
    def findMedianSortedArrays(self, nums1, nums2) -> float:
        if len(nums1) > len(nums2):
            nums1, nums2 = nums2, nums1
        length1 = len(nums1)
        length2 = len(nums2)
        k = (length1 + length2 + 1) // 2
        l, r = -1, length1
        while l + 1 != r:
            m = (l+r) // 2
            n = k - (m+1) - 1
            if nums1[m] <= nums2[n+1] and nums2[n] <= nums1[m+1]:
                if (length1 + length2) % 2 == 0:
                    return (max(nums1[m], nums2[n]) + min(nums1[m+1], nums2[n+1])) / 2
                else:
                    return max(nums1[m], nums2[n])
            elif nums1[m] > nums2[n+1]:
                r = m
            elif nums2[n] > nums1[m+1]:
                l = m

if __name__ == '__main__':
    nums1 = [1,7,8,10,15,18,21, 25]
    nums2 = [2,11,15,20,30,40]
    solution = Solution()
    print(solution.findMedianSortedArrays(nums1, nums2))