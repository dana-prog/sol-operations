const _eventProps = {
  sheetName: 'Events',
  name: 'event',
  itemDialog: {
    height: 450,
    width: 400,
  },
  fields: {
    type: {
      name: 'type',
      columnHeader: 'Type',
      required: true,
      showInForm: true,
      inputProps: {
        type: 'select',
        options: [
          {value: 'Retreat', label: 'Retreat'},
          {value: 'Parents Workshop', label: 'Parents Workshop'},
        ]
      }
    },
    sequence: {
      name: 'sequence',
      columnHeader: 'Sequence',
      required: true,
      showInForm: true,
      inputProps: {
        type: 'number',
      }
    },
    startDate: {
      name: 'startDate',
      columnHeader: 'Start Date',
      showInForm: true,
      inputProps: {
        type: 'date',
      }
    },
    endDate: {
      name: 'endDate',
      columnHeader: 'End Date',
      showInForm: true,
      inputProps: {
        type: 'date',
      }
    },
    fullPrice: {
      name: 'fullPrice',
      columnHeader: 'Full Price',
      showInForm: true,
      inputProps: {
        type: 'number',
      }
    },
    discountedPrice: {
      name: 'discountedPrice',
      columnHeader: 'Discounted Price',
      showInForm: true,
      inputProps: {
        type: 'number',
      }
    },
  }
};

addTypeProps(_eventProps);
