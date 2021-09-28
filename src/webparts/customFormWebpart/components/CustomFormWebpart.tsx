import * as React from 'react';
import styles from './CustomFormWebpart.module.scss';
import { ICustomFormWebpartProps } from './ICustomFormWebpartProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { ICustomFormWebpartState } from './ICustomFormWebpartState';
import { PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { ListItemPicker } from '@pnp/spfx-controls-react/lib/ListItemPicker';
import { ListPicker } from '@pnp/spfx-controls-react/lib/ListPicker';

import { sp } from "@pnp/sp";

import { Label, TextField, ChoiceGroup, Checkbox, IChoiceGroupOption, Button } from 'office-ui-fabric-react';

export default class CustomFormWebpart extends React.Component<ICustomFormWebpartProps, ICustomFormWebpartState> {

  constructor(props: ICustomFormWebpartProps) {
    super(props);
    this.getEmail = this.getEmail.bind(this);
    this.getMobile = this.getMobile.bind(this);
    this.submitData = this.submitData.bind(this);
    sp.setup({
      spfxContext: this.props.context
    });
    this.state = {
      email: "",
      mobile: "",
      SeletedList: ""
    }
    this.onListPickerChange = this.onListPickerChange.bind(this);

    
  }

  public trainingType: IChoiceGroupOption[] = [{ key: 'Yes', text: 'Yes' }, { key: 'No', text: 'No' }];

  public Cancel() {

  }

  public submitData() {
    let validation: boolean = true;
    if (this.state.email == "") {
      validation = false;
      document.getElementById('validation_email').setAttribute("style", "display:block !important")
    } else {
      document.getElementById('validation_mobile').setAttribute("style", "display:none !important")
    }

    if (this.state.mobile == "") {
      validation = false;
      document.getElementById('validation_mobile').setAttribute("style", "display:block !important")
    } else {
      document.getElementById('validation_mobile').setAttribute("style", "display:none !important")
    }


    if (validation) {

    }
  }

  private onListPickerChange(selectedlist: string) {
    this.setState({
      SeletedList: selectedlist
    });

  }
  private onSelectedItem(data: { key: string; name: string }[]) {
    for (const item of data) {
      console.log(`Item value: ${item.key}`);
      console.log(`Item text: ${item.name}`);
    }
  }
  private getEmail(newValue: string): void {
    this.setState({
      email: newValue
    })
  }

  private getMobile(newValue: string): void {
    this.setState({
      mobile: newValue
    })
  }

  public render(): React.ReactElement<ICustomFormWebpartProps> {
    return (
      <div className={styles.customFormWebpart} >
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.title}>Trainig Request Form!</div>


            <div id="custForm">
              <div className={styles.grid}>
                <div className={styles.gridRow}>

                  <div className={styles.rowDiv}>
                    <div className={styles.smallCol}>
                      <Label>Employee Name<span className={styles.validation}>*</span></Label>

                    </div>
                    <div className={styles.largeCol}>
                      <PeoplePicker ensureUser={true} personSelectionLimit={1} context={this.props.context} placeholder={'Enter id of Employee'} principalTypes={[PrincipalType.User]}></PeoplePicker>
                    </div>
                  </div>

                  <div className={styles.rowDiv}>
                    <div className={styles.smallCol}>
                      <Label>Email</Label>
                    </div>
                    <div className={styles.largeCol}>
                      <TextField onChanged={this.getEmail} placeholder="Enter your email" />
                      <div id="validation_email" className="form_validation">
                        <span>you can't leave this blank</span>
                      </div>
                    </div>
                  </div>

                  <div className={styles.rowDiv}>
                    <div className={styles.smallCol}>
                      <Label>Mobile No</Label>
                    </div>
                    <div className={styles.largeCol}>
                      <TextField onChanged={this.getMobile} type="number" placeholder="Enter your mobile no." />
                      <div id="validation_mobile" className="form_validation">
                        <span>you can't leave this blank</span>
                      </div>
                    </div>
                  </div>

                  <div className={styles.rowDiv}>
                    <div className={styles.smallCol}>
                      <Label>Address<span className={styles.validation}>*</span></Label>
                    </div>
                    <div className={styles.largeCol}>
                      <TextField multiline={true} />
                    </div>
                  </div>

                  <div className={styles.rowDiv}>
                    <div className={styles.smallCol}>
                      <Label>Do you have manager approval?<span className={styles.validation}>*</span></Label>
                    </div>
                    <div className={styles.largeCol}>
                      <ChoiceGroup options={this.trainingType}></ChoiceGroup>
                    </div>
                  </div>

                  <div className={styles.rowDiv}>
                    <div className={styles.smallCol}>
                      <Label>Are you available on weekend?<span className={styles.validation}>*</span></Label>
                    </div>
                    <div className={styles.largeCol}>
                      <Checkbox label="Yes"></Checkbox>
                    </div>
                  </div>

                  <div className={styles.rowDiv}>
                    <div className={styles.smallCol}>
                      <Label>Are you available on weekend?<span className={styles.validation}>*</span></Label>
                    </div>
                    <div className={styles.largeCol}>
                      <ListPicker context={this.props.context}
                        label="Select your list"
                        placeHolder="Select your list"
                        baseTemplate={100}
                        includeHidden={false}
                        multiSelect={false}
                        onSelectionChanged={this.onListPickerChange} />
                      <br></br>
                      <label>Search List Item</label>
                      <div>  

                      <ListItemPicker listId={this.state.SeletedList}
                        columnInternalName='Title'
                        itemLimit={5}
                        onSelectedItem={this.onSelectedItem}
                        context={this.props.context} />
                        </div>
                    </div>
                  </div>

                  {/* <div className={styles.rowDiv}>
                    <div className={styles.smallCol}>
                      <Label>Are you available on weekend?<span className={styles.validation}>*</span></Label>
                    </div>
                    <div className={styles.largeCol}>
                      <FieldCollectionData
                        key={"FieldCollectionData"}
                        label={"Fields Collection"}
                        manageBtnLabel={"Manage"} onChanged={(value) => { console.log(value); }}
                        panelHeader={"Manage values"}

                        executeFiltering={(searchFilter: string, item: any) => {
                          return item["Field2"] === +searchFilter;
                        }}
                        itemsPerPage={3}
                        fields={[
                          { id: "Field1", title: "String field", type: CustomCollectionFieldType.string, required: true },
                          { id: "Field2", title: "Number field", type: CustomCollectionFieldType.number },
                          { id: "Field3", title: "URL field", type: CustomCollectionFieldType.url },
                          { id: "Field4", title: "Boolean field", type: CustomCollectionFieldType.boolean }
                        ]}
                        value={[
                          {
                            "Field1": "String value", "Field2": "123", "Field3": "https://pnp.github.io/", "Field4": true
                          }
                        ]}
                      />
                    </div>
                  </div> */}



                  <div className={styles.largeCol}>
                    <Button className={styles.button} text={'Submit'} onClick={this.submitData}></Button>
                    <Button className={styles.button} text={'Cancel'} onClick={this.Cancel}></Button>

                  </div>


                </div>
              </div>
            </div>

          </div>
        </div>
      </div >
    );
  }
}
