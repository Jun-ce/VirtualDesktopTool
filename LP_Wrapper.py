# -*- coding: utf-8 -*-

import wrapt
from line_profiler import LineProfiler

lp = LineProfiler()

def _lp_wrapper():
    @wrapt.decorator
    def wrapper(func, instance, args, kwargs):
        global lp
        lp_wrapper = lp(func)
        res = lp_wrapper(*args, **kwargs)
        lp.print_stats()
        return res

    return wrapper

lp_wrapper = _lp_wrapper()

# @lp_wrapper
# def demo():
#     # 执行的函数
#     i = 1
#     print(i)

# # demo()