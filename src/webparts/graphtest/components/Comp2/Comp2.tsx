import * as React from 'react';
import {TextField} from '@fluentui/react/lib/TextField';

import { IComp2Props } from './IComp2Props';
import { Checkbox, DefaultButton, Dropdown, Label, PrimaryButton, SpinButton, Toggle,IDropdownOption, ComboBox } from '@fluentui/react';




// import { version } from 'react-dom';

const dropDownDepartment:IDropdownOption[] = [
  {key:'IT',text:'IT'},
  {key:'HR',text:'HR'}
]


export default class Graphtest extends React.Component<IComp2Props,{}> {






  public render(): React.ReactElement<IComp2Props> {
    

    return (
      <div>
        <h1>Welcome to Component2</h1>
        <div>
          
        <TextField label="Username" prefix="UN"/>
        </div>
        <div>
          
        <TextField label="Paasword" type='password' canRevealPassword/>
        </div>

        <PrimaryButton type=''></PrimaryButton>
        <TextField type='text' multiline
     rows={5}
     iconProps={{iconName:'home'}}
     />

     <PrimaryButton text='Trash' iconProps={{iconName:'delete'}}/>
     <DefaultButton text='Send' iconProps={{iconName:'send'}}/>
     <SpinButton min={1} max={1000}/>

     <div>
      <Toggle  onText='ON' offText='OFF'/>
      <Toggle onText='ON' offText='OFF' defaultChecked/>
     </div>

     <Label>CheckBox</Label>
 
     <Checkbox label='India'/>

     <Label>Dropdown:</Label>
<Dropdown
options={[
  {key:'IT',text:'IT'},
  {key:'Finance',text:'Finance'},
  {key:'Audit',text:'Audit'}
 
]}
placeholder='select an options'
/>

<Label> Multi Select Dropdown:</Label>
<Dropdown
options={[
  {key:'IT',text:'IT'},
  {key:'Finance',text:'Finance'},
  {key:'Audit',text:'Audit'}
 
]}
placeholder='select an options'
multiSelect
/>

<Dropdown options={dropDownDepartment} defaultSelectedKey={['IT','HR']} multiSelect/>



<Label>ComboBox:</Label>
<ComboBox
options={[
  {key:'IT',text:'IT'},
  {key:'Finance',text:'Finance'},
  {key:'Audit',text:'Audit'}
 
]}allowFreeform
autoComplete='on'
multiSelect/>


    </div>
    );
  }
}
