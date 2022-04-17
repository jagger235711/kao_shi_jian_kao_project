"""
 -*- coding: utf-8 -*-

 @Time : 2021/11/9 19:46

 @Author : jagger

 @File : test.py

 @Software: PyCharm 

 @contact: 252587809@qq.com

 -*- 功能说明 -*-

"""


def quickSort(arr, left=None, right=None):
    '''
    快速排序的递归实现
    Parameters
    ----------
    arr : 待排序数组
    left : 首个元素
    right : 最后一个元素

    Returns
    -------
    arr 结果数组
    '''
    left = 0 if not isinstance(left, (int, float)) else left
    right = len(arr) - 1 if not isinstance(right, (int, float)) else right
    if left < right:
        partitionIndex = partition(arr, left, right)
        quickSort(arr, left, partitionIndex - 1)
        quickSort(arr, partitionIndex + 1, right)
    return arr


def partition(arr, left, right):
    '''
    一趟划分
    Parameters
    ----------
    arr : 待排序数组
    left : 首个元素
    right : 最后一个元素

    Returns
    -------

    '''
    pivot = left  # 枢轴
    index = pivot + 1
    i = index
    while i <= right:
        if arr[i] < arr[pivot]:
            swap(arr, i, index)
            index += 1
        i += 1
    swap(arr, pivot, index - 1)
    return index - 1


def swap(arr, i, j):
    '''
    交换函数
    Parameters
    ----------
    arr : 目标数组
    i : 元素下标
    j : 元素下标

    Returns
    -------

    '''
    arr[i], arr[j] = arr[j], arr[i]


if __name__ == "__main__":
    alist = [54, 26, 93, 17, 77, 31, 44, 55, 20]
    print(quickSort(alist))
