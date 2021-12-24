def days_between(d1, d2) -> int:
    """
    This function returns the absolute difference between two data.
    :param d1: Closing data
    :param d2: Opening data
    :return: Days between
    """
    return abs((d2 - d1).days)