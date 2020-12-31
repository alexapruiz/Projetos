========
dataview
========

**dataview** is a module that allows you to create views of your sequences or its slices

Install
-------

.. code-block::

    pip install dataview

To upgrade a previous installation, use:

.. code-block::

    pip install dataview

Usage && Examples
-----------------

.. code-block::

    >>> from dataview import DataView
    >>> # You have some data, it can be any sequence (str, list, bytes, tuple, etc..)
    >>> source_data = list(range(5))
    >>> # DataView is just a pointer to your source_data and start/stop/step
    >>> DataView(source_data)
    [0, 1, 2, 3, 4]
    >>> DataView(source_data, 3)
    [0, 1, 2]
    >>> DataView(source_data, 1, 5)
    [1, 2, 3, 4]
    >>> DataView(source_data, None, None, -1)
    [4, 3, 2, 1, 0]
    >>> # You can use slices (completely the same way as list slices)
    >>> DataView(source_data)[::-1]
    [4, 3, 2, 1, 0]
    >>> # Slice return a new DataView object, that points to the previous DataView
    >>> DataView(source_data)[::-1][2:4]
    [2, 1]
    >>> # You can change start/stop/step anytime
    >>> view = DataView(source_data, 0, 1)
    >>> view.start = 1
    >>> view.stop = 2
    >>> view
    [1]
    >>> # View always points to actual data
    >>> source_data[1] = 2
    >>> view
    [2]
    >>> # You can change a view source data
    >>> view.data = list(range(5))
    >>> view
    [1]










