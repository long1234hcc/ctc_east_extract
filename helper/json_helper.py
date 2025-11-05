import numpy as np

def convert_numpy_types(obj):
    if isinstance(obj, dict):
        return {convert_numpy_types(key): convert_numpy_types(value) for key, value in obj.items()}
    elif isinstance(obj, list):
        return [convert_numpy_types(item) for item in obj]
    elif isinstance(obj, (np.integer, np.int64)):
        return str(obj)
    elif isinstance(obj, (np.floating, np.float64)):
        return str(obj)
    elif isinstance(obj, np.ndarray):
        return obj.tolist()
    return obj