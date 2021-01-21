from mailmerge import MailMerge
import os
import csv
import logging
import subprocess
g_logger = logging.getLogger(__name__)

class CSVReader:

    def __init__(self, csv_file_name):
        self.__file_name = csv_file_name
    def __call__(self):
        name = []
        email = []

        with open(self.__file_name, 'r') as csv_file:
            csv_reader = csv.reader(csv_file)

            next(csv_reader)

            for line in csv_reader:
                name.append(line[0])
                email.append(line[1])


        if len(name) != len(email):
            g_logger.error(" file %s is invalid/corrupted", self.__file_name)
        _details = {}
        for i in range(len(name)):
            _details[name[i]] = email[i]
        g_logger.error("_details:\n{}".format(_details))
        return _details

class Create_Docx_Pdf:
	def __init__(self, template, name, email):
		self.__template = template
		self.__name = name
		self.__email = email
		self.__new_offer_letter = ""

	def __call__(self):

		g_logger.debug("Name: %s", self.__name)
		g_logger.debug("Address: %s", self.__email)
		g_logger.debug("Offer Letter Template Name: {}".format(self.__template))

		name = self.__name
		file = '{}.docx'.format(name)

		with MailMerge(self.__template) as document:
			document.merge(fieldname = name)
			document.write(file)

		if os.path.exists(file):
			subprocess.run(['abiword', '--to=pdf', file])


def main():
    global g_logger
    logging.basicConfig(filename="Send_Offer_Letter.log", filemode='w', level=logging.DEBUG, )
    g_logger.info("Staring the script")
    details = CSVReader('test1.csv')()
    g_logger.debug("main() - details: {}".format(details))
    for name in details.keys():
        offer_letter = Create_Docx_Pdf("test_merge_pages.docx", name, details[name])()

if __name__ == '__main__':
	main()