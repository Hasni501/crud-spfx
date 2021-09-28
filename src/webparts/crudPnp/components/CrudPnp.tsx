import * as React from 'react';
import styles from './CrudPnp.module.scss';
import { ICrudPnpProps } from './ICrudPnpProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { ICrudPnpState } from './ICrudPnpState';
import { SPOperations } from '../../Services/SPServices';
import { Button, Dropdown, IDropdownOption } from 'office-ui-fabric-react';


export default class CrudPnp extends React.Component<ICrudPnpProps, ICrudPnpState> {
  public _SPOps: SPOperations;

  public selectedListTitle: string;

  constructor(props: ICrudPnpProps) {
    super(props);
    this.state = {
      title: 'arsalan1111',
      listTitles: [],
      status: ''
    }
    this._SPOps = new SPOperations();

  }

  public getListTitle(event: any, result: any) {
    this.selectedListTitle = event.text;
  }

  public componentDidMount() {

    this._SPOps.getAllListPNP().then((result) => {
      this.setState({
        listTitles: result
      })
    })

  }

  public render(): React.ReactElement<ICrudPnpProps> {
    return (
      <div className={styles.crudPnp} >
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>Welcome to SharePoint!</span>
              <p className={styles.subTitle}>Customize SharePoint experiences using Web Parts.</p>

            </div>
            <div className={styles.myStyles}>
              <Dropdown className={styles.dropdow} options={this.state.listTitles} onChanged={(e, selectedItem) => this.getListTitle(e, selectedItem)}
              ></Dropdown>
              <Button className={styles.myButton} text="Create List Item" onClick={() => this._SPOps.createListItePNP(this.selectedListTitle).then((result: string) => {
                this.setState({
                  status: result
                })
              })}></Button>
             <Button  className={styles.myButton} text="Delete List Item" onClick={() => this._SPOps.deleteListItePNP(this.selectedListTitle).then((result: string) => {
                this.setState({
                  status: result
                })
              })}></Button>

            <Button  className={styles.myButton} text="Update List Item" onClick={() => this._SPOps.updateListItePNP(this.selectedListTitle).then((result: string) => {
                this.setState({
                  status: result
                })
              })}></Button>
            </div>

            <div  className={styles.myStatusBar}>{this.state.status}</div>
          </div>
        </div>
      </div >
    );
  }
}
