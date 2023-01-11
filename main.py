from events import Events


def main():
    filename = 'capex.xlm'
    events = Events(filename)
    print(events.chapter8_summary[8])


if __name__ == '__main__':
    main()
