#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
@project: axislib
@file: python-class-wrapper.py
@author: modeldevtools
@date: 10/19/2020
"""
import types


class Wrapper(object):
	def __init__(self, obj):
		self._obj = obj

	def __getattr__(self, attr):

		if hasattr(self._obj, attr):
			attr_value = getattr(self._obj, attr)

			if isinstance(attr_value, types.MethodType):
				def callable(*args, **kwargs):
					return attr_value(*args, **kwargs)

				return callable
			else:
				return attr_value

		else:
			raise AttributeError
	
if __name__ == '__main__':
	class AddMethod(Wrapper):
		def bar(self):
			print
			"Call 'bar' method"


	class A(object):
		def __init__(self, value):
			self.value = value

		def foo(self, f1, f2="str"):
			print
			"Call: ", self.foo, " args: (", f1, f2, ")"


	wrapped = AddMethod(A("object value"))

	wrapped.foo(2, 3)
	wrapped.bar()
	print(wrapped.value)
	
