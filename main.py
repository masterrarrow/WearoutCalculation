import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from decimal import Decimal


"""
    Create plots (wear out of bushings and pins) and calculate wear out difference
"""


def main():
    # Load data for bushings
    bushing = pd.DataFrame(pd.read_excel('wearout.xls', 'Bushing'))

    # Load data for pins
    pin = pd.DataFrame(pd.read_excel('wearout.xls', 'Pin'))

    """
        Calculate wear out difference
    """

    # Wear out difference for bushings
    start = bushing['bushing_1, gr.'][0]
    lst1 = []
    for v in bushing['bushing_1, gr.']:
        lst1.append(format(Decimal.from_float(start - v), '.1'))
        start = v

    start = bushing['bushing_2, gr.'][0]
    lst2 = []
    for v in bushing['bushing_2, gr.']:
        lst2.append(format(Decimal.from_float(start - v), '.1'))
        start = v

    start = bushing['bushing_3, gr.'][0]
    lst3 = []
    for v in bushing['bushing_3, gr.']:
        lst3.append(format(Decimal.from_float(start - v), '.1'))
        start = v

    wearout_bushing = pd.DataFrame({
        'bushing_1, gr.': lst1,
        'bushing_2, gr.': lst2,
        'bushing_3, gr.': lst3,
        'time, min': bushing['range, min']
    })

    # Wear out difference for pins
    start = pin['pin_1, gr.'][0]
    lst1 = []
    for v in pin['pin_1, gr.']:
        lst1.append(format(Decimal.from_float(start - v), '.1'))
        start = v

    start = pin['pin_2, gr.'][0]
    lst2 = []
    for v in pin['pin_2, gr.']:
        lst2.append(format(Decimal.from_float(start - v), '.1'))
        start = v

    start = pin['pin_3, gr.'][0]
    lst3 = []
    for v in pin['pin_3, gr.']:
        lst3.append(format(Decimal.from_float(start - v), '.1'))
        start = v

    wearout_pin = pd.DataFrame({
        'pin_1, gr.': lst1,
        'pin_2, gr.': lst2,
        'pin_3, gr.': lst3,
        'time, min': pin['range, min']
    })

    """
        Create plots (wear out of bushings and pins)
    """

    fig = plt.figure(figsize=(6, 9))

    # Polot wear out of bushings
    bushings_plot = fig.add_subplot(211)

    # Settings
    bushings_plot.set_title('Wear out of bushings and pins')
    bushings_plot.set_xlabel('Time, minutes')
    bushings_plot.set_ylabel('Wear out, grams')

    labels = list(x for x in range(0, 50, 5))
    bushings_plot.set_xticks(labels)

    label = list(x for x in np.arange(97.6, 98.2, 0.05))
    bushings_plot.set_yticks(label)

    bushings_plot.grid()

    bushings_plot.set_xlim(-1, 46)
    bushings_plot.set_ylim(97.6, 98.1)

    # Set plots
    bushings_plot.plot(
        bushing['range, min'], bushing['bushing_1, gr.'], marker='D', label='Bushing 1')
    bushings_plot.plot(
        bushing['range, min'], bushing['bushing_2, gr.'], marker='D', label='Bushing 2')
    bushings_plot.plot(
        bushing['range, min'], bushing['bushing_3, gr.'], marker='D', label='Bushing 3')

    bushings_plot.legend(shadow=True)

    # Polot wear out of pins
    pin_plot = fig.add_subplot(212)

    # Settings
    pin_plot.set_xlabel('Time, minutes')
    pin_plot.set_ylabel('Wear out, grams')

    labels = list(x for x in range(0, 55, 5))
    pin_plot.set_xticks(labels)

    label = list(x for x in np.arange(66.6, 67.15, 0.05))
    pin_plot.set_yticks(label)

    pin_plot.grid()

    pin_plot.set_xlim(-1, 46)
    pin_plot.set_ylim(66.6, 67.1)

    # Set plots
    pin_plot.plot(pin['range, min'], pin['pin_1, gr.'],
                  marker='D', label='Pin 1')
    pin_plot.plot(pin['range, min'], pin['pin_2, gr.'],
                  marker='D', label='Pin 2')
    pin_plot.plot(pin['range, min'], pin['pin_3, gr.'],
                  marker='D', label='Pin 3')

    pin_plot.legend(shadow=True)

    # Save plot
    plt.savefig('wearout.png', bbox_inches='tight',
                pad_inches=0, transparent=False)

    """
        Save data and plot to Excel
    """

    # Write data into Excel file
    with pd.ExcelWriter('wearout_data.xls', engine='xlsxwriter') as writer:
        # Write wear out
        bushing.to_excel(writer, sheet_name='Wearout', index=False, startrow=1)
        pin.to_excel(writer, sheet_name='Wearout', index=False,
                     startrow=(len(bushing['range, min']) + 3))

        # Write wear out difference
        wearout_bushing.to_excel(
            writer, sheet_name='Wearout_difference', index=False, startrow=1)
        wearout_pin.to_excel(writer, sheet_name='Wearout_difference', index=False,
                             startrow=(len(wearout_bushing['time, min']) + 3))

        # Write an image (plot)
        worksheet = writer.sheets['Wearout']
        worksheet.insert_image('F2', 'wearout.png')

    plt.show()


if __name__ == '__main__':
    main()
