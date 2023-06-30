import functools
import time

__all__ = [
    "profile"
]

# Used to track execution time of designated functions
def profile(func):
    """
    Print the runtime of the decorated function.
    """
    @functools.wraps(func)
    def wrapper_profile(*args, **kwargs):
        start_time = time.perf_counter()    # 1
        value = func(*args, **kwargs)
        end_time = time.perf_counter()      # 2
        run_time = end_time - start_time    # 3
        print(f"- exec {func.__name__!r}: {run_time:.4f}s")
        return value
    return wrapper_profile