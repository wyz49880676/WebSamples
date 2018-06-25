import * as React from 'react';
import styles from './RecentDocs.module.scss';
import { IDocItemProps, IDocItemState, Direction } from '../Model/IRecentDocsProps';
import SPProvider from '../Controller/SPClientProvider'

export default class DocItem extends React.Component<IDocItemProps, IDocItemState> {
    constructor(props: IDocItemProps, state: IDocItemState) {
        super(props);
        this.state = {
            icon: `${this.props.context.pageContext.web.absoluteUrl}/_layouts/15/images/lg_icgen.gif`,
            isLoaded: false
        }
    }

    public render(): React.ReactElement<IDocItemProps> {

        if (!this.state.isLoaded) {
            const spProvider = new SPProvider(this.props.context);
            spProvider.LoadDocIcon(this.props.title).then((data: string) => {
                this.setState({
                    icon: data,
                    isLoaded: true,
                });
            }, (error: any) => {
                this.setState({
                    icon: `${this.props.context.pageContext.web.absoluteUrl}/_layouts/15/images/lg_icgen.gif`,
                    isLoaded: true,
                });
            });
        }

        let width;
        if (this.props.direct == Direction.vertical) {
            width = '100%';
        }
        else {
            width = '260px';
        }

        return (
            <li className={styles.item}>
                <table style={{ 'width': width }}>
                    <tbody>
                        <tr>
                            <td className={styles.icon}>
                                <img src={this.state.icon}></img>
                            </td>
                            <td className={styles.doc}>
                                <table>
                                    <tbody>
                                        <tr>
                                            <td>
                                                <a href={this.props.url}>{this.props.title}</a>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td>
                                                <span>Modified Date: {this.props.modifyDate}</span>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td>
                                                <span>By: {this.props.modifyBy}</span>
                                            </td>
                                        </tr>
                                    </tbody>
                                </table>
                            </td>
                        </tr>
                    </tbody>
                </table>
            </li>
        );
    }
    private onWindowResize(e) {
        this.setState(this.state);
    }
    public componentDidMount(): void {
        window.addEventListener('resize', this.onWindowResize.bind(this));
    }
    public componentWillUnmount(): void {
        window.removeEventListener('resize', this.onWindowResize.bind(this));
    }
}