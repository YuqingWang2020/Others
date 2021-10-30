"""
Y.Wang
Date: 30.10.2021
"""
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# @Time    : 2019/1/21 21:06
# @Author  : Arrow and Bullet
# @FileName: gradient_descent.py
# @Software: PyCharm
# @Blog    ：https://blog.csdn.net/qq_41800366

import numpy as np

# 数据集大小 即20个数据点
m = 20
# x的坐标以及对应的矩阵
X0 = np.ones((m, 1))  # 生成一个m行1列的向量，也就是x0，全是1
#print (X0)
X1 = np.arange(1, m+1)  #[ 1  2  3  4  5  6  7  8  9 10 11 12 13 14 15 16 17 18 19 20] #arange返回一个array对象
X1 = X1.reshape(m, 1) # 生成一个m行1列的向量，也就是x1，从1到m
#print (X1)
X = np.hstack((X0, X1))  # 按照列堆叠形成数组，其实就是样本数据 #hstack用于将两数组或矩阵合并
#print (X)
# 对应的y坐标
'''
a=[
    3, 4, 5, 5, 2, 4, 7, 8, 11, 8, 12,
    11, 13, 13, 16, 17, 18, 17, 19, 21
]
print(type(a)) #list
'''
Y = np.array([
    3, 4, 5, 5, 2, 4, 7, 8, 11, 8, 12,
    11, 13, 13, 16, 17, 18, 17, 19, 21
]).reshape(m, 1)
#print(y)
# 学习率
alpha = 0.01

# 定义代价函数
def cost_function(theta, X, Y):
    diff = np.dot(X, theta) - Y  # dot() 数组需要像矩阵那样相乘，就需要用到dot()
    return (1/(2*m)) * np.dot(diff.transpose(), diff) #transpose()函数的作用就是调换数组的行列值的索引值，类似于求矩阵的转置

# 定义代价函数对应的梯度函数
def gradient_function(theta, X, Y):
    diff = np.dot(X, theta) - Y
    return (1/m) * np.dot(X.transpose(), diff)

# 梯度下降迭代
def gradient_descent(X, Y, alpha):
    theta = np.array([1, 1]).reshape(2, 1)
    gradient = gradient_function(theta, X, Y)
    while not all(abs(gradient) <= 1e-5):   #Python all（）函数是内置函数之一。 它以iterable作为参数，如果iterable的所有元素均为true或为空，则返回True
        theta = theta - alpha * gradient
        gradient = gradient_function(theta, X, Y)
    return theta
#print(np.array([1, 1])) #[1 1]

optimal = gradient_descent(X, Y, alpha)
print('optimal:', optimal)  # 结果 [[0.51583286][0.96992163]]
print('cost function:', cost_function(optimal, X, Y)[0][0])  # 1.014962406233101


# 根据数据画出对应的图像
def plot(X, Y, theta):
    import matplotlib.pyplot as plt
    ax = plt.subplot(111)  # #1代表行，1代表列，所以一共有1个图，1代表此时绘制第二个图。其中ax是为了坐标轴主次刻度大小的设置
    ax.scatter(X, Y, s=30, c="red", marker="s")
    plt.xlabel("X")
    plt.ylabel("Y")
    x = np.arange(0, 21, 0.2)  # x的范围
    y = theta[0] + theta[1]*x
    ax.plot(x, y)
    plt.show()


plot(X1, Y, optimal)

