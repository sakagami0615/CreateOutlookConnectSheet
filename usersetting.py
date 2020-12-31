# 表示名生成関数
# 名字、名前、アドレスからOutlook表示名を作成する
# 今回は、[名字+名前+ドメイン]となるようにしている
def GetDisplayName(lastname, firstname, address):

	tokens = address.split('@')
	if len(tokens) > 1:
		domain = tokens[len(tokens) - 1]

		if domain == 'gmail.com':
			return '{}{}@GMail'.format(lastname, firstname)
		
		elif domain == 'yahoo.co.jp':
			return '{}{}YahooMail'.format(lastname, firstname)
	
	return '{}{}'.format(lastname, firstname)

# エクセル連絡先シートのヘッダ名
# (今回はテスト用の値を格納している)
# 予め日本語 or 英語のヘッダ名を格納しておく必要がある
OUTLOOK_FORMAT_HEADER = ['lastname', 'firstname', 'displayname', 'address']


# 引数の値をエクセル連絡先シートのレコードに変換する関数
# エクセル連絡先シート用の辞書を作成する
def ConvertOutlookFormat(lastname, firstname, displayname, address):
	
	outlook_connect = dict()

	# 空文字で初期化する
	for key in OUTLOOK_FORMAT_HEADER: outlook_connect[key] = ''

	# OUTLOOK_FORMAT_HEADERの文字列をキーにして値を格納する
	outlook_connect['lastname'] = lastname
	outlook_connect['firstname'] = firstname
	outlook_connect['displayname'] = displayname
	outlook_connect['address'] = address

	return outlook_connect
