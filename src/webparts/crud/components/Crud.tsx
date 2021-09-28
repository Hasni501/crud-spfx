import * as React from 'react';
import styles from './Crud.module.scss';
import { ICrudProps } from './ICrudProps';
import { ICrudState } from './ICrudState';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPOperations } from '../../Services/SPServices';
import { Button, Dropdown, IDropdownOption } from 'office-ui-fabric-react';

export default class Crud extends React.Component<ICrudProps, ICrudState>  {

  public _SPOps: SPOperations;

  public selectedListTitle: string;

  constructor(props: ICrudProps) {
    super(props);
    this.state = {
      title: 'arsalan1111',
      listTitle: [],
      status: ''
    };
    this._SPOps = new SPOperations();
    this.testMethod = this.testMethod.bind(this);
  }

  test = (hi: any) => {
    alert('hi' + hi)
    return 'test'
  }

  public testMethod() {
    this.setState({
      title: 'arsalan'
    })
    alert('ayajghg');
  }

  testMethod2 = (val: string) => {
    this.setState({
      title: 'arsalan'
    })
    alert(val);

  }

  public getListTitle = (event: any, data: any) => {
    this.selectedListTitle = event.text;
  }

  public getListTitles = () => {
    alert('hi')
  }


  public componentDidMount() {
    this._SPOps.getAllList(this.props.context).then((result: IDropdownOption[]) => {
      this.setState({
        listTitle: result
      })
    });
  }

  public render(): React.ReactElement<ICrudProps> {

    return (
      <div className={styles.crud} >
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>Welcome to SharePoint!</span>
              <p className={styles.subTitle}>Customize SharePoint experiences using Web Parts.</p>
              
              <div>
                {/* <button onClick={this.testMethod}>{this.state.title}</button>
                <button onClick={() => this.testMethod2('arsalan here')}>{this.state.title}</button> */}

              </div>
              <div className={styles.myStyles}>
                <Dropdown className={styles.dropdow} options={this.state.listTitle}                
                onChanged={(e, selectedItem) => this.getListTitle(e, selectedItem)}
                placeHolder="***Select List Title***" ></Dropdown>
                <Button className={styles.myButton} text="Create List Item" onClick={() => this._SPOps.createListTitle(this.props.context, this.selectedListTitle).then((result: string) => {
                  this.setState({
                    status: result
                  })
                })}></Button>
                <Button className={styles.myButton} onClick={()=>this._SPOps.updateListItem(this.props.context,this.selectedListTitle).then((result:any)=>{
                  this.setState({
                    status:result
                  })
                })} text="Update List Item" ></Button>
                <Button className={styles.myButton} onClick={()=>this._SPOps.deleteListItem(this.props.context,this.selectedListTitle).then((result:any)=>{
                  this.setState({
                    status:result
                  })
                })} text="Delete List Item" ></Button>
                <div className={styles.myStatusBar}>{this.state.status}</div>
              </div>
            </div>
          </div>
        </div>
      </div>

    );
  }
}
