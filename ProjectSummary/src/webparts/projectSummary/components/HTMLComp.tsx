import * as React from 'react';
import { Label, TextField } from 'office-ui-fabric-react/lib';
import styles from './ProjectSummary.module.scss';
export default class HTMLContent extends React.Component {
    render() {
        return (
            <div>
                <div className={styles.row}>
                    <TextField label="Investors/Partners" placeholder="List of Investors/Partners" required />
                </div>
            </div>
        )

    }
}

