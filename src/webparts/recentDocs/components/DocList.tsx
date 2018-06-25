import * as React from 'react';
import styles from './RecentDocs.module.scss';
import { IDocListProps, IDocListState, IDocItemProps, Direction } from '../Model/IRecentDocsProps';
import DocItem from './DocItem';
import SPProvider from '../Controller/SPClientProvider';
import { Spinner } from 'office-ui-fabric-react/lib/Spinner';

export default class DocList extends React.Component<IDocListProps, IDocListState> {
    private icon = "https://spoprod-a.akamaihd.net/files/odsp-next-prod_2018-06-08-sts_20180613.001/odsp-media/images/itemtypes/20/folder.svg";

    constructor(props: IDocListProps, state: IDocListState) {
        super(props);

        this.state = {
            items: [],
            isLoaded: false,
            message: null
        }
    }

    public render(): React.ReactElement<IDocListProps> {
        const loadingElement: JSX.Element = <div style={{ 'margin-top': '40%' }}><Spinner label={'Loading Recent Documents...'} /></div>;
        const errorElement: JSX.Element = <div className={'ms-TextField-errorMessage ms-u-slideDownIn20'}>Error while loading items: {this.state.message}</div>;

        if (!this.state.isLoaded) {
            const spProvider = new SPProvider(this.props.context);
            spProvider.LoadLatestDocs(this.props.pageCount).then((data: IDocItemProps[]) => {
                let count = 0;

                let docItems = data.map((item: any) => {
                    count++;
                    let id = 'doc_item_' + count;
                    return <DocItem key={id}
                        context={this.props.context}
                        title={item.title}
                        modifyDate={item.modifyDate}
                        modifyBy={item.modifyBy}
                        url={item.url} 
                        direct={this.props.direct} />
                });

                this.setState({
                    items: docItems,
                    isLoaded: true,
                    message: null
                });

            }, (error: any) => {
                this.setState({
                    items: [],
                    isLoaded: true,
                    message: error
                });
            });

            if (this.state.message != null) {
                return errorElement;
            }
            else {
                return loadingElement;
            }
        }

        return (
            <ul className={styles.list} style={{ 'display': 'block' }}>
                {this.state.items}
            </ul>
        );
    }

    private onWindowResize(e) {
        this.setState({
            items: this.state.items,
            isLoaded: false,
            message: null
        });
    }
    public componentDidMount(): void {
        window.addEventListener('resize', this.onWindowResize.bind(this));
    }
    public componentWillUnmount(): void {
        window.removeEventListener('resize', this.onWindowResize.bind(this));
    }
}