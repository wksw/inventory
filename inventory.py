# -*- encoding: utf-8 -*-

# 按店主
# {
# 	"name": {
# 		"goods_list": [{
# 			"name": "",
# 			"flavor": "",
# 			"number": 0
# 		}],
# 		"pay_status": "",
# 		"phone": "",
# 		"address": "",
# 	}
# }

# 按商品
# {
# 	"goods_name": {
# 		"flavors":[{
# 			"bussiness_name": ""
# 			"flavor": "",
# 			"number": ""
# 		}],
# 	}
# }


import csv
import xlwt
import re
import sys
import os

ROOTDIR = os.path.dirname(os.path.realpath(__file__))
OUTPUT=os.path.join(ROOTDIR, "output")

def write_output(output, output1, filename):
	wb = xlwt.Workbook(encoding="utf-8")
	sheet1 = wb.add_sheet("发货单", cell_overwrite_ok = True)
	sheet2 = wb.add_sheet("采购单", cell_overwrite_ok = True)
	# 设置列宽
	sheet1.col(0).width = 3333 * 1
	sheet1.col(1).width = 3333 * 2
	sheet1.col(2).width = 3000
	sheet1.col(3).width = 1666
	sheet1.col(4).width = 2000
	sheet1.col(5).width = 2000
	sheet1.col(6).width = 2000
	sheet1.col(7).width = 2000
	sheet1.col(8).width = 3333 * 1
	sheet1.col(9).width = 3333
	sheet1.col(10).width = 3333 * 3

	sheet2.col(0).width = 3333 * 2
	sheet2.col(1).width = 3333 * 1
	sheet2.col(2).width = 1666
	sheet2.col(3).width = 1666
	sheet2.col(4).width = 1666
	sheet2.col(5).width = 1666
	sheet2.col(6).width = 3333

	# 头部颜色，边框
	pattern = xlwt.Pattern()
	pattern.pattern = xlwt.Pattern.SOLID_PATTERN
	pattern.pattern_fore_colour = 7

	# 边框
	borders = xlwt.Borders()
	borders.left = xlwt.Borders.HAIR
	borders.right = xlwt.Borders.HAIR
	borders.top = xlwt.Borders.HAIR
	borders.bottom = xlwt.Borders.HAIR

	# 居中
	alignment = xlwt.Alignment()
	alignment.horz = xlwt.Alignment.HORZ_CENTER
	alignment.vert = xlwt.Alignment.VERT_CENTER

	header_style = xlwt.XFStyle()
	header_style.pattern = pattern
	header_style.borders = borders
	header_style.alignment = alignment

	center_style = xlwt.XFStyle()
	center_style.alignment = alignment

	# 颜色区分
	pattern1 = xlwt.Pattern()
	pattern1.pattern = xlwt.Pattern.SOLID_PATTERN
	pattern1.pattern_fore_colour = 22
	color_style = xlwt.XFStyle()
	color_style.pattern = pattern1
	color_style.borders = borders
	color_style.alignment = alignment


	
	header = ["店主", "商品名", "规格", "数量", "实际斤数", "店铺价格", "单件价格", "总额（含运费）", "付款状态", "电话", "地址"]
	goods_list_len = 0
	keys = output.keys()
	for key_index in range(len(output)):
		key = keys[key_index]
		# 头部内容写入
		for index in range(0, len(header)):
			sheet1.write(key_index + goods_list_len, index, header[index], header_style)
		# 商品
		for goods_index in range(0, len(output[key]["goods_list"])):
			# 商品名
			sheet1.write(goods_index + key_index+goods_list_len + 1, 1, output[key]["goods_list"][goods_index]["name"].decode("gbk"), center_style)
			# 商品规格
			sheet1.write(goods_index + key_index+goods_list_len + 1, 2, output[key]["goods_list"][goods_index]["flavor"].decode("gbk"), center_style)
			# 商品数量
			sheet1.write(goods_index + key_index+goods_list_len + 1, 3, output[key]["goods_list"][goods_index]["number"].decode("gbk"), center_style)

		# 店主
		sheet1.write_merge(key_index+goods_list_len+1, key_index+goods_list_len + 1 + len(output[key]["goods_list"]) - 1, 0, 0, keys[key_index].decode("gbk"), center_style)
		# 付款状态
		sheet1.write_merge(key_index+goods_list_len+1, key_index+goods_list_len + 1 + len(output[key]["goods_list"]) - 1, 8, 8, output[key]["pay_status"].decode("gbk"), center_style)
		# 电话
		sheet1.write_merge(key_index+goods_list_len+1, key_index+goods_list_len + 1 + len(output[key]["goods_list"]) - 1, 9, 9, output[key]["phone"], center_style)
		# 地址
		sheet1.write_merge(key_index+goods_list_len+1, key_index+goods_list_len + 1 + len(output[key]["goods_list"]) - 1, 10, 10, output[key]["address"].decode("gbk"), center_style)
		goods_list_len += len(output[key]["goods_list"]) + 1


	header1 = ["商品名", "规格", "数量", "进货斤数", "进货价", "对外售价", "购货人", "总规格", "总数量"]
	flavor_list_len = 0
	# 头部写入
	for index in range(len(header1)):
		sheet2.write(0, index, header1[index], header_style)
	keys = output1.keys()
	for key_index in range(1, len(output1)):
		style = center_style
		if key_index % 2 == 0:
			style = color_style
		key = keys[key_index]
		# 总规格
		total_flavor = 0
		# 总数量
		total_number = 0
		# 规格单位
		flavor_unit = ""
		# 数量单位
		number_unit = ""
		for flavor_index in range(0, len(output1[key]["flavors"])):
			flavor_re = re.findall(r'\d+', output1[key]["flavors"][flavor_index]["flavor"].decode("gbk"))
			if flavor_re:
				total_flavor += int(flavor_re[-1])
			number_re = re.findall(r'\d+', output1[key]["flavors"][flavor_index]["number"].decode("gbk"))
			if number_re:
				total_number += int(number_re[-1])

			flavor_unit_re = re.findall(r'.*\d+(.+?),', output1[key]["flavors"][flavor_index]["flavor"].decode("gbk"))
			if flavor_unit_re:
				flavor_unit = flavor_unit_re[-1]
			number_unit_re = re.findall(r'.*\d+(.+?)$', output1[key]["flavors"][flavor_index]["number"].decode("gbk"))
			if number_unit_re:
				number_unit = number_unit_re[-1]

			# 规格
			sheet2.write(flavor_index + flavor_list_len + 1, 1, output1[key]["flavors"][flavor_index]["flavor"].decode("gbk"), style)
			# 数量
			sheet2.write(flavor_index + flavor_list_len + 1, 2, output1[key]["flavors"][flavor_index]["number"].decode("gbk"), style)
			# 进货斤数
			sheet2.write(flavor_index + flavor_list_len + 1, 3, "", style)
			# 进货价
			sheet2.write(flavor_index + flavor_list_len + 1, 4, "", style)
			# 对外售价
			sheet2.write(flavor_index + flavor_list_len + 1, 5, "", style)
			# 购货人
			sheet2.write(flavor_index + flavor_list_len + 1, 6, output1[key]["flavors"][flavor_index]["bussiness_name"].decode("gbk"), style)


		# 商品名
		sheet2.write_merge(flavor_list_len + 1, flavor_list_len + len(output1[key]["flavors"]), 0, 0, key.decode("gbk"), style)
		# 总规格
		sheet2.write_merge(flavor_list_len + 1, flavor_list_len + len(output1[key]["flavors"]), 7, 7, str(total_flavor) + flavor_unit, style)
		# 数量
		sheet2.write_merge(flavor_list_len + 1, flavor_list_len + len(output1[key]["flavors"]), 8, 8, str(total_number) + number_unit, style)
		flavor_list_len += len(output1[key]["flavors"])

	wb.save(os.path.join(OUTPUT, filename + "_output.xls"))

def useage():
	print '''
使用方法：
	python inventory.py <文件名>
		文件名： 需要解析的CSV文件(必须为CSV文件)
		''' 
	sys.exit(1)
if __name__ == "__main__":
	if len(sys.argv) < 2:
		useage()
	# 判断文件是否存在
	if not os.path.isfile(sys.argv[1]):
		useage()
	# 创建输出目录

	if not os.path.isdir(OUTPUT):
		os.mkdir(OUTPUT)
	filename = os.path.splitext(os.path.basename(sys.argv[1]))[0]
	print "正在解析", sys.argv[1], "..."
	with open(sys.argv[1]) as file:
		csv_file = csv.reader(file)
		header = next(csv_file)
		# 按店主解析
		output1 = {}
		# 按商品解析
		output2 = {}
		for line in csv_file:
			if not output1.has_key(line[6]):
				output1[line[6]] = {}
				output1[line[6]]["goods_list"] = []
			output1[line[6]]["goods_list"].append({
				"name": line[3],
				"flavor": line[4],
				"number": line[5]
				})
			output1[line[6]]["pay_status"] = line[17]
			output1[line[6]]["phone"] = line[7]
			output1[line[6]]["address"] = line[8]

			# 按商品解析
			if not output2.has_key(line[3]):
				output2[line[3]] = {}
				output2[line[3]]["flavors"] = []
			output2[line[3]]["flavors"].append({
					"bussiness_name": line[6],
					"flavor": line[4],
					"number": line[5]
				})
		
		write_output(output1, output2, filename)
		print "解析完成"		
