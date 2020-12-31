import unicodedata


def CalcStringWidth(string_value):
	
	def GetCharWidth(char_value):
		char_kind = unicodedata.east_asian_width(char_value)
		if (char_kind == 'H') or (char_kind == 'Na'):
			char_width = 1
		else:
			char_width = 2
		return char_width
	
	char_widths = [GetCharWidth(char_value) for char_value in string_value]
	
	return sum(char_widths)
