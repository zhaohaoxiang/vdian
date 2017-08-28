#!/usr/local/bin/python3.6
# -*- coding: utf-8 -*-     

##
## Description:
## 	
##	
##
## Author: Harry Zhao


import re
import os
import sys
import shutil


class Order:
	def __init__(self, id, name, phone, num, province, city, district, addr, note, pay_date):
		self.id = id
		self.name = name
		self.phone = phone
		self.num = num
		self.province = province
		self.city = city
		self.district = district
		self.addr = addr
		self.note = note
		self.pay_date = pay_date
		
		
class Product:
	def __init__(self, p_name, price):
		self.p_name = p_name
		self.price = price
		self.order_num = 0
		self.sell_num = 0
		self.orders = []		


class Vdian:
	def __init__(self):
		self.products = []
	
	def add_order(self, p_name, price, id, name, phone, num, province, city, district, addr, note, pay_date):
		order = Order(id, name, phone, num, province, city, district, addr, note, pay_date)
		found = False
		for p in self.products:
			if p.p_name == p_name:
				found = True
				break
		if found:
			product = p
		else:
			product = Product(p_name, price)
			self.products.append(product)
		found = False
		for o in product.orders:
			if o.id == id:
				print("[WARNING] Duplicate order!!! " + id + " - " + name + " - " + str(phone) + " - " + p_name)
				found = True
				break
		if not found:
			product.orders.append(order)
			product.order_num = product.order_num + 1
			product.sell_num = product.sell_num + int(num)
	
	def show(self):
		for p in self.products:
			print("\n=== Product ===")
			print(p.p_name)
			print("订单数量: " + str(p.order_num))
			print("商品数量: " + str(p.sell_num))
		
		
		