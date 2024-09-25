from parsing.config import Config

GUI = Config().gui.gui

if not GUI:
    from not_gui.wood_waste_calc_not_gui import main
else:
    from gui.wood_waste_calc_gui import main

if __name__ == '__main__':
    main()
