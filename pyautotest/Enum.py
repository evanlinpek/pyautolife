# coding = utf-8
from enum import Enum, unique

@unique
class Browser(Enum):

    CHROME = 1
    EDGE = 2
    FIREFOX = 3
    IE = 4
    SAFARI = 5



if __name__ == '__main__':

    browser = Browser(Browser.EDGE)
    print(browser)
    print(Browser.EDGE.value)