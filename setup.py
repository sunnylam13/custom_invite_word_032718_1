try:
	from setuptools import setup
except ImportError:
	from distutils.core import setup

config = {
	'description': 'Creates a custom invitation Word document using a list of names in a text file.',
	'author': 'Sunny Lam',
	'url': 'URL to get it at',
	'download_url': 'Where to download it',
	'author_email': 'sunny.lam@gmail.com',
	'version': '0.1',
	'install_requires': ['nose'],
	'packages': ['docx'],
	'scripts': [],
	'name': 'Custom Invitations as Word Documents'
}

setup(**config)