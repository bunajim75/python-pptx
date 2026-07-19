.. _install:

Installing
==========

|pp| is hosted on PyPI, so installing with `pip` is simple::

    pip install python-pptx

|pp| depends on the ``lxml`` package and ``Pillow``, the modern version of
the Python Imaging Library (``PIL``). The charting features depend on
``XlsxWriter``. Both ``pip`` and ``easy_install`` will take care of
satisfying these dependencies for you, but if you use the ``setup.py``
installation method you will need to install the dependencies yourself.

Currently |pp| requires Python 3.11 or later. The tests are run against Python 3.11
and 3.12 in CI.

Dependencies
------------

* Python 3.11 or later
* lxml
* Pillow
* XlsxWriter (to use charting features)
