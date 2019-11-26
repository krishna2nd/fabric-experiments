import * as React from 'react';
import { Link } from 'office-ui-fabric-react/lib/Link';
import {
  DetailsList,
  Selection,
  IColumn,
  buildColumns,
  IColumnReorderOptions,
  IDragDropEvents,
  IDragDropContext
} from 'office-ui-fabric-react/lib/DetailsList';
import { MarqueeSelection } from 'office-ui-fabric-react/lib/MarqueeSelection';
import { createListItems, IExampleItem } from '@uifabric/example-data';
import { TextField, ITextFieldStyles } from 'office-ui-fabric-react/lib/TextField';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { getTheme, mergeStyles } from 'office-ui-fabric-react/lib/Styling';
import { values } from '@uifabric/utilities';

const theme = getTheme();
const margin = '0 30px 20px 0';
const dragEnterClass = mergeStyles({
  backgroundColor: theme.palette.neutralLight
});
const controlWrapperClass = mergeStyles({
  display: 'flex',
  flexWrap: 'wrap'
});
const textFieldStyles: Partial<ITextFieldStyles> = {
  root: { margin: margin },
  fieldGroup: { maxWidth: '100px' }
};

export interface IExampleItemX extends IExampleItem {
    index: number
};

export interface IDetailsListDragDropExampleState {
  items: IExampleItemX[];
  columns: IColumn[];
  isColumnReorderEnabled: boolean | undefined;
  frozenColumnCountFromStart: string | undefined;
  frozenColumnCountFromEnd: string | undefined;
}

const placeHolder = {
    key: 'placeholder'
} as IExampleItemX;

export class DetailsListDragDropExample extends React.Component<{}, IDetailsListDragDropExampleState> {
  private _selection: Selection;
  private _dragDropEvents: IDragDropEvents;
  private _draggedItem: IExampleItemX | undefined;
  private _draggedIndex: number;

  constructor(props: {}) {
    super(props);

    this._selection = new Selection();
    this._dragDropEvents = this._getDragDropEvents();
    this._draggedIndex = -1;
    let items = createListItems(10, 0);
    let index = 0;
    const itemsX = items.map((item: IExampleItem): IExampleItemX => {
        const itemX = item as IExampleItemX;
        itemX.index = index++;
        return itemX;
    })
    this.state = {
      items: itemsX,
      columns: buildColumns(items, true),
      isColumnReorderEnabled: true,
      frozenColumnCountFromStart: '1',
      frozenColumnCountFromEnd: '0'
    };
  }


  public render(): JSX.Element {
    const { items, columns, isColumnReorderEnabled, frozenColumnCountFromStart, frozenColumnCountFromEnd } = this.state;

    return (
      <div>
        <div className={controlWrapperClass}>
          <Toggle
            label="Enable column reorder"
            checked={isColumnReorderEnabled}
            onChange={this._onChangeColumnReorderEnabled}
            onText="Enabled"
            offText="Disabled"
            styles={{ root: { margin: margin } }}
          />
          <TextField
            label="Number of left frozen columns"
            onGetErrorMessage={this._validateNumber}
            value={frozenColumnCountFromStart}
            onChange={this._onChangeStartCountText}
            styles={textFieldStyles}
          />
          <TextField
            label="Number of right frozen columns"
            onGetErrorMessage={this._validateNumber}
            value={frozenColumnCountFromEnd}
            onChange={this._onChangeEndCountText}
            styles={textFieldStyles}
          />
        </div>
        <MarqueeSelection selection={this._selection}>
          <DetailsList
            setKey="items"
            items={items}
            columns={columns}
            selection={this._selection}
            selectionPreservedOnEmptyClick={true}
            onItemInvoked={this._onItemInvoked}
            onRenderItemColumn={this._onRenderItemColumn}
            dragDropEvents={this._dragDropEvents}
            columnReorderOptions={this.state.isColumnReorderEnabled ? this._getColumnReorderOptions() : undefined}
            ariaLabelForSelectionColumn="Toggle selection"
            ariaLabelForSelectAllCheckbox="Toggle selection for all items"
            checkButtonAriaLabel="Row checkbox"
          />
        </MarqueeSelection>
      </div>
    );
  }

  private _handleColumnReorder = (draggedIndex: number, targetIndex: number) => {
    const draggedItems = this.state.columns[draggedIndex];
    const newColumns: IColumn[] = [...this.state.columns];

    // insert before the dropped item
    newColumns.splice(draggedIndex, 1);
    newColumns.splice(targetIndex, 0, draggedItems);
    this.setState({ columns: newColumns });
  };

  private _getColumnReorderOptions(): IColumnReorderOptions {
    return {
      frozenColumnCountFromStart: parseInt(this.state.frozenColumnCountFromStart as string, 10),
      frozenColumnCountFromEnd: parseInt(this.state.frozenColumnCountFromEnd as string, 10),
      handleColumnReorder: this._handleColumnReorder
    };
  }

  private _validateNumber(value: string): string {
    return isNaN(Number(value)) ? `The value should be a number, actual is ${value}.` : '';
  }

  private _onChangeStartCountText = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text: string | undefined): void => {
    this.setState({ frozenColumnCountFromStart: text });
  };

  private _onChangeEndCountText = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text: string | undefined ): void => {
    this.setState({ frozenColumnCountFromEnd: text });
  };

  private _onChangeColumnReorderEnabled = (ev: React.MouseEvent<HTMLElement>, checked: boolean| undefined): void => {
    this.setState({ isColumnReorderEnabled: checked });
  };

  private _getDragDropEvents(): IDragDropEvents {
    return {
      canDrop: (dropContext?: IDragDropContext, dragContext?: IDragDropContext) => {
        return true;
      },
      canDrag: (item?: any) => {
        return true;
      },
      onDragEnter: (item?: any, event?: DragEvent) => {
        const items = this.state.items.reduce((values ,value) => {
            if (value.key !== 'placeholder') values.push(value as IExampleItemX); 
            return values
        } , [] as IExampleItemX[])
        let index = 0;
        const preItems = items.slice(0, item.index), postItems = items.slice(item.index, items.length);
        placeHolder.index = item.index;
        const newItems = [
            ...preItems,
            placeHolder,
            ...postItems
        ].map(_item => {
            _item.index = index++;
            return _item;
        });
        console.log(items.length, preItems.length, postItems.length, newItems);
        this.setState({
            items: newItems
        })
        // return string is the css classes that will be added to the entering element.
        return dragEnterClass;
      },
      onDragLeave: (item?: any, event?: DragEvent) => {
        let items = this.state.items;
        items = items.reduce((values ,value) => {
            if (value.key !== 'placeholder') values.push(value as IExampleItemX); 
            return values
        } , [] as IExampleItemX[])
        this.setState({
            items
        })
        return;
      },
      onDrop: (item?: any, event?: DragEvent) => {
        if (this._draggedItem) {
          this._insertBeforeItem(item);
        }
      },
      onDragStart: (item?: any, itemIndex?: number, selectedItems?: any[], event?: MouseEvent) => {
        this._draggedItem = item;
        this._draggedIndex = itemIndex!;
      },
      onDragEnd: (item?: any, event?: DragEvent) => {
        this._draggedItem = undefined;
        this._draggedIndex = -1;
      }
    };
  }

  private _onItemInvoked = (item: IExampleItemX): void => {
    alert(`Item invoked: ${item.name}`);
  };

  private _onRenderItemColumn = (item: IExampleItemX, index: number | undefined, column: IColumn | undefined): JSX.Element | string => {
    const key = (column && column.key || "") as keyof IExampleItemX;
    if (key === 'name') {
      return <Link data-selection-invoke={true}>{item[key]}</Link>;
    }

    return item && String(item[key]);
  };

  private _insertBeforeItem(item: IExampleItemX): void {
    const draggedItems = this._selection.isIndexSelected(this._draggedIndex)
      ? (this._selection.getSelection() as IExampleItemX[])
      : [this._draggedItem!];

    const items = this.state.items.filter(itm => draggedItems.indexOf(itm) === -1);
    let insertIndex = items.indexOf(item);

    // if dragging/dropping on itself, index will be 0.
    if (insertIndex === -1) {
      insertIndex = 0;
    }

    items.splice(insertIndex, 0, ...draggedItems);

    this.setState({ items: items });
  }
}
