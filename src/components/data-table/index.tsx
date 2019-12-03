import * as React from "react";
import { Link } from "office-ui-fabric-react/lib/Link";
import {
  DetailsList,
  Selection,
  IColumn,
  buildColumns,
  IColumnReorderOptions,
  IDragDropEvents,
  IDragDropContext
} from "office-ui-fabric-react/lib/DetailsList";
import { MarqueeSelection } from "office-ui-fabric-react/lib/MarqueeSelection";
import { createListItems, IExampleItem } from "@uifabric/example-data";
import {
  TextField,
  ITextFieldStyles
} from "office-ui-fabric-react/lib/TextField";
import { Toggle } from "office-ui-fabric-react/lib/Toggle";
import { getTheme, mergeStyles } from "office-ui-fabric-react/lib/Styling";

const theme = getTheme();
const margin = "0 30px 20px 0";
const dragEnterClass = mergeStyles({
  backgroundColor: theme.palette.neutralLight
});
const controlWrapperClass = mergeStyles({
  display: "flex",
  flexWrap: "wrap"
});
const textFieldStyles: Partial<ITextFieldStyles> = {
  root: { margin: margin },
  fieldGroup: { maxWidth: "100px" }
};

enum Direction {
  UP,
  DOWN
}

export interface IDetailsListDragDropExampleState {
  items: IExampleItem[];
  columns: IColumn[];
  isColumnReorderEnabled: boolean | undefined;
  frozenColumnCountFromStart: string | undefined;
  frozenColumnCountFromEnd: string | undefined;
}

export class DetailsListDragDropExample extends React.Component<
  {},
  IDetailsListDragDropExampleState
> {
  private _selection: Selection;
  private _dragDropEvents: IDragDropEvents;
  private _draggedItem: HTMLElement | null;
  private _draggedItemClone: HTMLElement | null;
  private _draggedOverItem: HTMLElement | null;
  private _pointerMoveDirection: Direction;

  constructor(props: {}) {
    super(props);

    this._selection = new Selection();
    this._dragDropEvents = this._getDragDropEvents();
    this._draggedItem = null;
    this._draggedItemClone = null;
    this._draggedOverItem = null;
    this._pointerMoveDirection = Direction.DOWN;
    let items = createListItems(10, 0);
    this.state = {
      items: items,
      columns: buildColumns(items, true),
      isColumnReorderEnabled: true,
      frozenColumnCountFromStart: "1",
      frozenColumnCountFromEnd: "0"
    };
  }

  public render(): JSX.Element {
    const {
      items,
      columns,
      isColumnReorderEnabled,
      frozenColumnCountFromStart,
      frozenColumnCountFromEnd
    } = this.state;

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
            columnReorderOptions={
              this.state.isColumnReorderEnabled
                ? this._getColumnReorderOptions()
                : undefined
            }
            ariaLabelForSelectionColumn="Toggle selection"
            ariaLabelForSelectAllCheckbox="Toggle selection for all items"
            checkButtonAriaLabel="Row checkbox"
          />
        </MarqueeSelection>
      </div>
    );
  }

  private _handleColumnReorder = (
    draggedIndex: number,
    targetIndex: number
  ) => {
    const draggedItems = this.state.columns[draggedIndex];
    const newColumns: IColumn[] = [...this.state.columns];

    // insert before the dropped item
    newColumns.splice(draggedIndex, 1);
    newColumns.splice(targetIndex, 0, draggedItems);
    this.setState({ columns: newColumns });
  };

  private _getColumnReorderOptions(): IColumnReorderOptions {
    return {
      frozenColumnCountFromStart: parseInt(
        this.state.frozenColumnCountFromStart as string,
        10
      ),
      frozenColumnCountFromEnd: parseInt(
        this.state.frozenColumnCountFromEnd as string,
        10
      ),
      handleColumnReorder: this._handleColumnReorder
    };
  }

  private _validateNumber(value: string): string {
    return isNaN(Number(value))
      ? `The value should be a number, actual is ${value}.`
      : "";
  }

  private _onChangeStartCountText = (
    event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    text: string | undefined
  ): void => {
    this.setState({ frozenColumnCountFromStart: text });
  };

  private _onChangeEndCountText = (
    event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    text: string | undefined
  ): void => {
    this.setState({ frozenColumnCountFromEnd: text });
  };

  private _onChangeColumnReorderEnabled = (
    ev: React.MouseEvent<HTMLElement>,
    checked: boolean | undefined
  ): void => {
    this.setState({ isColumnReorderEnabled: checked });
  };

  private _getDragDropEvents(): IDragDropEvents {
    return {
      canDrop: (
        _?: IDragDropContext,
        _1?: IDragDropContext
      ) => {
        return true;
      },
      canDrag: (_?: any) => {
        return true;
      },
      onDragEnter: (_?: any, event?: DragEvent) => {
        console.log("onDragEnter");
        const draggedItem = this._draggedItem;
        const draggedItemClone = this._draggedItemClone;
        if (draggedItem && draggedItem.style.display !== "none") {
          if (!draggedItem.parentElement) {
            console.error("draggedItem.parentElement in not available.");
          } else if (!draggedItemClone) {
            console.error("draggedItemClone");
          } else {
            draggedItem.style.display = "none";
            draggedItem.parentElement.insertBefore(
              draggedItemClone,
              draggedItem.nextElementSibling
            );
          }
        }

        const currentTarget = (event && event.currentTarget && (event.currentTarget as HTMLElement).parentElement) || null;
        if (event && currentTarget != draggedItem) {
          console.log("set draggedOverItem");
          this._draggedOverItem = currentTarget;
          this._pointerMoveDirection = getPointerDirection(
            this._draggedItemClone,
            event.pageY
          );

          if (this._pointerMoveDirection === Direction.DOWN) {
            if (!this._draggedOverItem) {
              console.error("this._draggedOverItem is not available");
            } else if (!this._draggedOverItem.parentElement) {
              console.error(
                "this._draggedOverItem.parentElement is not available"
              );
            } else if (!draggedItemClone) {
              console.error("draggedItemClone is not available");
            } else {
              this._draggedOverItem.parentElement.insertBefore(
                draggedItemClone,
                this._draggedOverItem.nextElementSibling
              );
            }
          } else {
            if (!this._draggedOverItem) {
              console.error("this._draggedOverItem is not available");
            } else if (!draggedItemClone) {
              console.error("draggedItemClone is not available");
            } else if (!this._draggedOverItem.parentElement) {
              console.error(
                "this._draggedOverItem.parentElement is not available"
              );
            } else {
              this._draggedOverItem.parentElement.insertBefore(
                draggedItemClone,
                this._draggedOverItem
              );
            }
          }
        }

        // return string is the css classes that will be added to the entering element.
        return dragEnterClass;
      },
      onDragLeave: (item?: any, event?: DragEvent) => {
        console.log("onDragLeave");
        console.log('event.currentTarget', event && event.currentTarget);
        if (event && event.currentTarget && (event.currentTarget as HTMLElement).parentNode === this._draggedOverItem) {
          this._draggedOverItem = null;
        }

        return;
      },
      onDragStart: (_?: any, _1?: number, _2?: any[], event?: MouseEvent) => {
        console.log("onDragStart");
        if (event && event.currentTarget instanceof HTMLElement) {
          //set draggedItem variable and a clone of it.
          this._draggedItem = event.currentTarget.parentElement;
          this._draggedItemClone = this._draggedItem && this._draggedItem.cloneNode(
            true
          ) as HTMLElement || null;
        }
      },
      onDragEnd: (_?: any, _1?: DragEvent) => {
        console.log("onDragEnd");
        if (this._draggedItem) {
          if (!this._draggedItem.parentElement) {
            console.error("this._draggedItem.parentElement is not available");
          } else if (this._draggedOverItem) {
            this._draggedItem.parentElement.removeChild(this._draggedItem);
          } else if (!this._draggedItemClone) {
            console.error("this._draggedItemClone is not available");
          } else {
            try {
              this._draggedItem.parentElement.removeChild(
                this._draggedItemClone
              );
              this._draggedItem.style.display = "";
            } catch (e) {
              console.log("clone not added to the DOM yet!");
            }
          }
          this._draggedItem = null;
          this._draggedItemClone = null;
        }
      }
    };
  }






















  private _onItemInvoked = (item: IExampleItem): void => {
    alert(`Item invoked: ${item.name}`);
  };

  private _onRenderItemColumn = (
    item: IExampleItem,
    index: number | undefined,
    column: IColumn | undefined
  ): JSX.Element | string => {
    const key = ((column && column.key) || "") as keyof IExampleItem;
    if (key === "name") {
      return <Link data-selection-invoke={true}>{item[key]}</Link>;
    }

    return item && String(item[key]);
  };
}

function getPointerDirection(
  draggedItemClone: HTMLElement | null,
  dragOverEventYPosition: number
): Direction {
  if (draggedItemClone) {
    return draggedItemClone.getBoundingClientRect().top < dragOverEventYPosition
      ? Direction.DOWN
      : Direction.UP;
  }
  return Direction.DOWN;
}
