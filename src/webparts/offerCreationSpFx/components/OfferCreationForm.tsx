import * as React from "react";
import { PrimaryButton } from '@fluentui/react/lib/Button';
import { DatePicker } from '@fluentui/react/lib/DatePicker';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { TextField } from '@fluentui/react/lib/TextField';
import { useState, useCallback } from "react";
import { IOfferCreationFormProps } from "./IOfferCreationFormProps";
import { IOffer } from "../../../model/IOffer";

export const OfferCreationForm = (props: IOfferCreationFormProps) => {
  const [title, setTitle] = useState<string>();
  const [date, setDate] = useState<string>();
  const [price, setPrice] = useState<string>();
  const [vat, setVAT] = useState<number>();
  const [selectedItem, setSelectedItem] = useState<IDropdownOption>();
  const [description, setTDescription] = useState<string>();
  
  const vatItems = [
    { key: '19', text: '19% (full)' },
    { key: '7', text: '7% (reduced)' }
  ];

  const onOfferingDateChange = (date: Date) => {
    setDate(date.toISOString());
  };

  const onOfferingVATChange = (e: React.FormEvent<HTMLDivElement>, selectedOption: IDropdownOption) => {
    setSelectedItem(selectedOption);
    switch (selectedOption.key) {
      case "19":
        setVAT(0.19);
        break;
      case "7":
        setVAT(0.07);
        break;
    }
  };

  const storeData = useCallback(() => {
    const newOffer: IOffer = {
      title: title ? title : '',
      description: description ? description : '',
      date: date ? date : '',
      price: price ? parseFloat(price) : 0,
      vat: vat ? vat : 0
    };
    props.createOffer(newOffer);
  }, [title, description, date, price,vat]);

  return (
    <div>
      <div>
        <TextField label="Title" 
                value={title}
                type="text" 
                onChange={(e, data) => {
                  if (data) {
                      setTitle(data);
                  }
                }} />          
      </div>
      <div>
        <DatePicker label="Offer Date" onSelectDate={onOfferingDateChange} />
      </div>
      <div>        
        <TextField label="Price" 
                value={price}
                type="number"
                step=".01"
                onChange={(e, data) => {
                  if (data) {
                    const numPrice = parseFloat(data);
                    if (!isNaN(numPrice)) {
                      setPrice(numPrice.toString());
                    }                    
                  }
                }} />          
      </div>
      <div>
        <Dropdown
          label="VAT"
          selectedKey={selectedItem ? selectedItem.key : undefined}
          // eslint-disable-next-line react/jsx-no-bind
          onChange={onOfferingVATChange}
          placeholder="Select VAT"
          options={vatItems}
        />
      </div>
      <div>
        <TextField label="Description" 
                    multiline rows={3} 
                    resizable 
                    value={description}
                    onChange={(e, data) => {
                      if (data) {
                          setTDescription(data);
                      }
                    }} />
      </div>
      <div>
        <PrimaryButton text="Create Offer" onClick={storeData} allowDisabledFocus />
      </div>
    </div>
  );
}