import * as React from 'react';
import { IAnniversaryWebPartProps } from './IAnniversaryWebPartProps';
import { IAnniversaryWebPartState } from './IAnniversaryWebPartState';
import { DetailsList, DetailsListLayoutMode, IColumn, IIconProps, IconButton, Link, SelectionMode } from 'office-ui-fabric-react';

// Import the necessary classes from PnPJS
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists/web";
import "@pnp/sp/views";
import "@pnp/sp/fields";
import "@pnp/sp/items";
import { getSP } from '../../../pnpjs-config';
import * as moment from 'moment';
import styles from './AnniversaryWebPart.module.scss';

export interface IDetailsListBasicExampleItem {
  key: number;
  name: string;
  date: string;
  dateSort: Date;
  milestone: string;
  email: string;
}

export default class AnniversaryWebPart extends React.Component<IAnniversaryWebPartProps, IAnniversaryWebPartState> {

  private _allItems: IDetailsListBasicExampleItem[];
  private _columns: IColumn[];
  private _sp: SPFI;
  private styleBlock: React.CSSProperties;
  //private timeoutId: number = -1;
  
  constructor(props: IAnniversaryWebPartProps) {
    super(props);

    this._sp = getSP();
    
    // Populate with items for demos.
    this._allItems = [];
    this.loadListItems();

    this._columns = [
      { key: 'name', name: 'Name', fieldName: 'name', minWidth: 100, maxWidth: 200, isResizable: true },
      { key: 'date', name: 'Date', fieldName: 'date', minWidth: 100, maxWidth: 200, isResizable: true },
      { key: 'milestone', name: 'Milestone', fieldName: 'milestone', minWidth: 100, maxWidth: 200, isResizable: true },
      { key: 'email', name: 'Email', fieldName: 'email', minWidth: 100, maxWidth: 200, isResizable: true },
    ];

    // Set initial state
    this.state = {
      isRunning: "running"
    }

  }

  public loadListItems() : void {

    // If we have a list selected
    if (this.props.listId) {

      const userField = this.props.nameFieldId;

      // Get the list items
      // TODO: Update to use the selected view
      // TODO: Update to only query if the name field is valid
      this._sp.web.lists.getById(this.props.listId).items.select(userField + "/EMail", 
                                                                 userField + "/FirstName", 
                                                                 userField + "/LastName", 
                                                                 this.props.dateFieldId, 
                                                                 this.props.milestoneFieldId)
                                                          .expand(userField)().then(items => {
          
        // Loop through the list items
        items.forEach((item, i) => {

          const name: string = `${item[userField].FirstName} ${item[userField].LastName}`;
          const date: Date = new Date(item[this.props.dateFieldId]);
          const milestone: string = item[this.props.milestoneFieldId];
          const email: string = `${item[userField].EMail}`;

          this._allItems.push({ 
            key: i,
            name: name,
            date: moment(date).format(this.props.dateFormat),
            dateSort: new Date(1980, date.getMonth(), date.getDate()),
            milestone: milestone,
            email: email
          });

        });

        this._allItems.sort((a: IDetailsListBasicExampleItem, b: IDetailsListBasicExampleItem) => a.dateSort < b.dateSort ? -1 : 1);
  
      })
      .catch(err => console.error(err));
  
    }
  }

  public render(): React.ReactElement<IAnniversaryWebPartProps> {
    
    const tileHeight: number = this.props.height;
    const speed:number = tileHeight / this.props.speed;
    const running:string = this.state.isRunning;

    this.styleBlock = { "--speed": speed + "s",
                       "--tileHeight": tileHeight + "px",
                       "--playState": running } as React.CSSProperties;

    return (
      <div className={ styles.container } style={ this.styleBlock }>
        <div className={ styles.wrapper } onMouseEnter={() => this.onMouseEnter()} onMouseLeave={() => this.onMouseOut()}>
          <DetailsList
              items={this._allItems}
              columns={this._columns}
              selectionMode={SelectionMode.none}
              setKey="set"
              layoutMode={DetailsListLayoutMode.justified}
              selectionPreservedOnEmptyClick={true}
              ariaLabelForSelectionColumn="Toggle selection"
              ariaLabelForSelectAllCheckbox="Toggle selection for all items"
              onRenderItemColumn={this.renderItemColumn}
              />
        </div>
      </div>
    );
  }

  public renderItemColumn(item: IDetailsListBasicExampleItem, index: number, column: IColumn): React.ReactNode {

    const fieldContent = item[column.fieldName as keyof IDetailsListBasicExampleItem] as string;
    const mailIcon: IIconProps = { iconName: 'Mail' };

    switch (column.key) {

      case 'email':
        return <Link href={"mailto:" + fieldContent}>
                 <IconButton iconProps={ mailIcon } title="Send Email" ariaLabel="Send Email" />
               </Link>;
        
      default:
        return <span>{fieldContent}</span>;
    }
  }

  public onMouseEnter(): void {

    // Pause the animation
    this.setState({isRunning: "paused"});

  }

  public onMouseOut(): void {
    
    // Resume the animation
    this.setState({isRunning: "running"});

  }

}
