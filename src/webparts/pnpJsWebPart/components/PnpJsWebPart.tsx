import * as React from 'react';
import type { IPnpJsWebPartProps } from './IPnpJsWebPartProps';
import { IPeoplePickerContext, PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { Label, PrimaryButton } from '@fluentui/react';
import { IPnpJsWebPartState } from './IPnpJsWebPartState';

const peoplePikcerInformation: IPnpJsWebPartState = {
  NameId: ''
}

interface IPeoplePickerServie {
  peoplePicker: IPnpJsWebPartState
}

const PnpJsWebPart: React.FC<IPnpJsWebPartProps> = (props: IPnpJsWebPartProps) => {

  const [value, setValue] = React.useState<IPnpJsWebPartState>(peoplePikcerInformation);

  const LIST_NAME = 'People Picker';

  const peoplePickerContext: IPeoplePickerContext = {
    absoluteUrl: props.context.pageContext.web.absoluteUrl,
    msGraphClientFactory: props.context.msGraphClientFactory,
    spHttpClient: props.context.spHttpClient
  };

  const _getPeoplePickerItems = (items: any[]) => {
    if (items.length > 0) {
      setValue({ ...value, NameId: items[0].id });
    }
  }
  const addPeoplePickerValue = async ({ peoplePicker = {} as IPnpJsWebPartState }: IPeoplePickerServie) => {
    let SiteURL: string = `${props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${LIST_NAME}')/items`;
    const headers: any = {
      'Accept': 'application/json;odata=verbose',
      'Content-Type': 'application/json; charset=utf-8',
      'odata-version': '',
    };
    const body: any = {
      "NameId": peoplePicker.NameId
    };
    try {
      const digestResponse = await fetch(`${props.context.pageContext.web.absoluteUrl}/_api/contextinfo`, {
        method: 'POST',
        headers: {
          'Accept': 'application/json;odata=verbose',
        }
      });
      const digestData = await digestResponse.json();
      headers['X-RequestDigest'] = digestData.d.GetContextWebInformation.FormDigestValue;
      const response = await fetch(SiteURL, {
        method: 'POST',
        headers: headers,
        body: JSON.stringify(body)
      });
      if (!response.ok) {
        throw new Error(`Error: ${response.statusText}`);
      }
      return response.json();
    } catch (error) {
      console.log('Something went wrong while adding the value of people picker ', error.message);
    }
  }

  const addPeoplePicker = async (event: React.FormEvent) => {
    event.preventDefault();
    try {
      const response = await addPeoplePickerValue({ peoplePicker: value });
      if (response) {
        alert('Name added successfully');
      }
    } catch (error) {
      console.log('Something went wrong addPeoplePicker ', error.message);
    }
  }
  console.log(value.NameId);
  return (
    <>
      <form onSubmit={addPeoplePicker}>
        <div>
          <Label>Select Person <span className='star'>*</span></Label>
          <PeoplePicker context={peoplePickerContext} placeholder='Enter name'
            personSelectionLimit={1}
            showtooltip={true}
            ensureUser={true}
            resolveDelay={100}
            principalTypes={[PrincipalType.User]}
            onChange={_getPeoplePickerItems} />
        </div>
        <div>
          <PrimaryButton type='submit' text='Submit' />
        </div>
      </form>

    </>
  );
}

export default PnpJsWebPart;
