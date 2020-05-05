
.. _shape_api:

Shape-related objects
=====================

.. currentmodule:: docx.shape


|InlineShapes| objects
----------------------

.. autoclass:: InlineShapes
   :members:
   :exclude-members: add_picture

|AnchorShapes| objects
----------------------

.. autoclass:: AnchorShapes
   :members:
   :exclude-members: add_background_picture

|InlineShape| objects
---------------------

The ``width`` and ``height`` property of |InlineShape| provide a length object
that is an instance of |Length|. These instances behave like an int, but also
have built-in units conversion properties, e.g.::

    >>> inline_shape.height
    914400
    >>> inline_shape.height.inches
    1.0

.. autoclass:: InlineShape
   :members: height, type, width

|AnchorShape| objects
---------------------

The ``width`` and ``height`` property of |AnchorShape| provide a length object
that is an instance of |Length|. These instances behave like an int, but also
have built-in units conversion properties, e.g.::

    >>> anchor_shape.height
    914400
    >>> anchor_shape.height.inches
    1.0

.. autoclass:: AnchorShape
   :members: height, type, width
