"""
Performance Wrapper for Excel Operations with Timeout Protection
"""
import time
import signal
from contextlib import contextmanager

class TimeoutException(Exception):
    pass

@contextmanager
def excel_read_timeout(seconds=5):
    """Context manager to enforce timeout on Excel read operations"""
    def timeout_handler(signum, frame):
        raise TimeoutException(f"Excel read exceeded {seconds}s timeout")
    
    # Note: signal.alarm only works on Unix. For Windows, we'll use threading
    import platform
    if platform.system() != 'Windows':
        old_handler = signal.signal(signal.SIGALRM, timeout_handler)
        signal.alarm(seconds)
        try:
            yield
        finally:
            signal.alarm(0)
            signal.signal(signal.SIGALRM, old_handler)
    else:
        # Windows fallback: just track time, no hard timeout
        start = time.time()
        yield
        elapsed = time.time() - start
        if elapsed > seconds:
            print(f"WARNING: Excel read took {elapsed:.2f}s (exceeds {seconds}s target)")

def timed_excel_read(file_path, operation_name="Excel Read", **kwargs):
    """
    Wrapper for pd.read_excel with timing and performance logging
    
    Args:
        file_path: Path to Excel file
        operation_name: Description for logging
        **kwargs: Arguments to pass to pd.read_excel
    
    Returns:
        DataFrame
    """
    import pandas as pd
    
    start_time = time.time()
    print(f"[PERF] Starting {operation_name}...")
    
    try:
        df = pd.read_excel(file_path, **kwargs)
        elapsed = time.time() - start_time
        
        rows, cols = df.shape
        print(f"[PERF] {operation_name} completed in {elapsed:.2f}s ({rows} rows, {cols} cols)")
        
        if elapsed > 3:
            print(f"[PERF] WARNING: {operation_name} took longer than 3s")
        
        return df
        
    except Exception as e:
        elapsed = time.time() - start_time
        print(f"[PERF] {operation_name} FAILED after {elapsed:.2f}s: {e}")
        raise
