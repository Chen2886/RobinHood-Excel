import six

if six.PY3:
    from RobinLib.Robinhood import Robinhood
else:
    from RobinLib import Robinhood
    import exceptions as RH_exception
