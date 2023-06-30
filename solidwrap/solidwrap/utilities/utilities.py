import functools
import time

# Import definition
__all__ = [
    "singleton",
    "profile"
]

# Used to enforce singleton behavior for designated classes
def singleton(cls):
    """
    Guards singleton class against multiple instantiation. New instances
    of the class will return always return the singleton instance.
    """
    @functools.wraps(cls)                                       # helper decorator to preserve class reference
    def wrapper_singleton(*args, **kwargs):                     # '*args, **kwargs' allows for infinte arguments
        if not wrapper_singleton.instance:                      # if an instance doesn't exist...
            wrapper_singleton.instance = cls(*args, **kwargs)   # ...then store instance
        return wrapper_singleton.instance                       # return instance
    wrapper_singleton.instance = None
    return wrapper_singleton
            

# Used to track execution time of designated processes
def profile(func):
    """
    Prints the runtime of the decorated function.
    """
    @functools.wraps(func)                  # helper decorator to preserve function reference
    def wrapper_profile(*args, **kwargs):   # '*args, **kwargs' allows for infinte arguments
        start_time = time.perf_counter()    # get time at start
        value = func(*args, **kwargs)
        end_time = time.perf_counter()      # get time at end
        run_time = end_time - start_time    # calculate runtime
        print(f"-- process {func.__name__!r} runtime: {run_time:.4f}s")
        return value
    return wrapper_profile