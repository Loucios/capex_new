from events import Events
from titles import Titles


def main():
    events = Events(Titles.filename)
    print(events.chapter12_summary[8])


if __name__ == '__main__':
    main()
