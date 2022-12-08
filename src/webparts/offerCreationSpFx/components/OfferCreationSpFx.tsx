import * as React from 'react';
import { useState } from "react";
import styles from './OfferCreationSpFx.module.scss';
import { IOfferCreationSpFxProps } from './IOfferCreationSpFxProps';
import { OfferCreationForm } from "./OfferCreationForm";
import { IOffer } from '../../../model/IOffer';
import { Label } from '@fluentui/react/lib/Label';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { SPService } from '../../../services/SPService';

export const OfferCreationSpFx: React.FC<IOfferCreationSpFxProps> = (props) => {
  const [offerCreated, setOfferCreated] = useState<boolean>(false);
  const [offerFileUrl, setOfferFileUrl] = useState<string>("");
  const [showSpinner, setShowSpinner] = useState<boolean>(false);

  const createOffer = (offer: IOffer) => {
    setShowSpinner(true);
    const _customSPServiceInstance = props.serviceScope.consume(SPService.serviceKey);

    _customSPServiceInstance.createOffer(offer, props.siteUrl, props.siteDomain).then((resp: any) => {
      console.log(resp);
      setOfferCreated(true);
      setShowSpinner(false);
      setOfferFileUrl(resp.fileUrl);
    })
    .catch((error) => {
        console.log(error);
        setShowSpinner(false);
    });
      // Axios.post(`https://${process.env.PUBLIC_HOSTNAME}/api/createoffer`, requestBody, {
      //             responseType: "json",
      //             headers: {
      //               Authorization: `Bearer ${idToken}`
      //             }
      // }).then(result => {
      //   if (result.data.fileUrl) {
      //     setOfferCreated(true);
      //   }     
      // })
      // .catch((error) => {
      //   console.log(error);
      // })
      // .finally(() => {
      //   setShowSpinner(false);
      // });
  };

  return (
    <section className={`${styles.offerCreationSpFx} ${props.hasTeamsContext ? styles.teams : ''}`}>
      <h2>Offer creation</h2>
      <OfferCreationForm offerCreated={offerCreated} createOffer={createOffer} />
      <div>{showSpinner && <Spinner label="Creating document" size={SpinnerSize.large} />}</div>
      <div>{offerCreated && <span><Label>Your offer document is created and can be found </Label><a href={offerFileUrl}>here</a></span>}</div>
    </section>
  );
}
