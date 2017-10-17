from drexel_class import Drexel
from rice_class import Rice


def main():
    drexel = Drexel()
    drexel.compct_xlsx_py('drexel/', 'output/drexel/')
    drexel.compct_xlsx_all('drexel/', 'output/drexel/')
    drexel.compct_xlsx_all_chinese('drexel/', 'output/drexel/')
    # rice = Rice()
    # rice.spider()


if __name__ == '__main__':
    main()
