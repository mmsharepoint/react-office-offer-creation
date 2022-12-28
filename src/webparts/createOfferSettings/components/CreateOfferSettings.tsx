import * as React from 'react';
import { useCallback, useState, useEffect } from "react";
import styles from './CreateOfferSettings.module.scss';
import { ICreateOfferSettingsProps } from './ICreateOfferSettingsProps';
import { TextField } from '@fluentui/react/lib/TextField';
import { PrimaryButton } from 'office-ui-fabric-react';
import GraphService from '../../../services/GraphService';

export const CreateOfferSettings: React.FC<ICreateOfferSettingsProps> = (props) => {
  const [siteUrl, setSiteUrl] = useState<string>();

  const storeData = useCallback(() => {
    const graphServiceInstance = props.serviceScope.consume(GraphService.serviceKey);
    graphServiceInstance.storePersonalSiteUrl(siteUrl)
    .catch((error) => {
      console.log(error);
    });
  }, [siteUrl]);

  useEffect((): void => {
    const graphServiceInstance = props.serviceScope.consume(GraphService.serviceKey);
    graphServiceInstance.getPersonalSiteUrl()
      .then(response => {
        setSiteUrl(response);
      });
  }, []);


  return (
    <section className={`${styles.createOfferSettings} ${props.hasTeamsContext ? styles.teams : ''}`}>
      <div className={styles.welcome}>
        <h2>Configuration</h2>
      </div>
      <div>
        <TextField label="Offer Site Url" 
                value={siteUrl}
                type="text" 
                onChange={(e, data) => {
                  if (data) {
                    setSiteUrl(data);
                  }
                }} />  
      </div>
      <div>
        <PrimaryButton text='Save' onClick={storeData} />
      </div>
    </section>
  );
}
