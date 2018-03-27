try:
	from setuptools import setup
except ImportError:
	from distutils.core import setup

config = {
	'description': 'Creates a custom invitation Word document using a list of names in a text file.',
	'author': 'Sunny Lam',
	'url': 'https://github.com/sunnylam13/custom_invite_word_032718_1',
	'download_url': 'https://github.com/sunnylam13/custom_invite_word_032718_1',
	'author_email': 'sunny.lam@gmail.com',
	'version': '0.1',
	'install_requires': ['nose'],
	'packages': ['docx'],
	'scripts': [],
	'name': 'Custom Invitations as Word Documents'
}

setup(**config)