import * as React from 'react';
import styles from './PeoplePickerWebpart.module.scss';
import { IPeoplePickerWebpartProps } from './IPeoplePickerWebpartProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { peoplePicker } from 'office-ui-fabric-react/lib/components/pickers/PeoplePicker/PeoplePicker.scss';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { PeoplePicker } from '@pnp/spfx-controls-react/lib/PeoplePicker'
import { Button } from 'office-ui-fabric-react';
import { IPeoplePickerWebpartState } from './IPeoplePickerWebpartState'
import { sp } from "@pnp/sp"

export default class PeoplePickerWebpart extends React.Component<IPeoplePickerWebpartProps, IPeoplePickerWebpartState> {

  constructor(props: IPeoplePickerWebpartProps) {
    super(props);
    this.state = {
      user: []
    }

    this.getPeoplePicker = this.getPeoplePicker.bind(this)

  }

  setPeoplePicker = () => {
    sp.web.lists.getByTitle('test list').items.add({ Title: 'People Picker Entry', EmployeeNameId: { results: this.state.user }, }).then(() => { alert('Successfully Submitted') })

  }

  public getPeoplePicker(items: any[]) {
    let itemsUser: any[] = [];

    items.map((item) => {
      itemsUser.push(item.id)
    })

    this.setState({
      user: itemsUser
    })
    console.log(items)

  }


  public render(): React.ReactElement<IPeoplePickerWebpartProps> {
    return (
      <div className={styles.peoplePickerWebpart} >
        <div className={styles.container}>
          <div className={styles.row}>
            <div id='div_custom'>
              <PeoplePicker ensureUser={true} selectedItems={this.getPeoplePicker} personSelectionLimit={1} context={this.props.context} titleText={'Employee Name'} placeholder={'Enter id of Employee'}></PeoplePicker>
              <Button text={'Set People Picker'} onClick={this.setPeoplePicker}></Button>
            </div>
          </div>
        </div>
      </div >
    );
  }
}
