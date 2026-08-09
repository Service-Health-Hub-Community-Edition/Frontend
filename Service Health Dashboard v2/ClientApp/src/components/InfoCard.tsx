import * as React from 'react';
import { Text } from '@fluentui/react';
import { FontIcon, mergeStyles, mergeStyleSets } from '@fluentui/react';
import './InfoCard.css'

interface IInfoCardState {
    icon?: string;
    iconColor: string;
    title?: string;
    href?: string;
    hrefTitle?: string;
    width?: number;
    height?: number;
}

export class InfoCard extends React.Component<{ icon?: string, iconColor?: string, title?: string, href?: string, hrefTitle?: string, width?: number, height?: number}, IInfoCardState> {
    constructor(props: { icon?: string, iconColor?: string, title?: string, href?: string, hrefTitle?: string, width?: number, height?: number}) {
        super(props);

        this.state = {
            icon: this.props.icon,
            iconColor: this.props.iconColor ? this.props.iconColor : "#2B2B2B",
            title: this.props.title,
            href: this.props.href,
            hrefTitle: this.props.hrefTitle,
            width: this.props.width ? this.props.width : 300,
            height: this.props.height ? this.props.height : 380
        };

    }

    public render() {
        const {
            icon, iconColor, title, href, hrefTitle, width, height
        } = this.state;

        const iconClass = mergeStyles({
            fontSize: 20,
            height: 20,
            width: 20,
            margin: '4px 0px 4px 4px',
            color: iconColor
        });

        const popExpandIconClass = mergeStyles({
            fontSize: 16,
            height: 16,
            width: 16,
            margin: '6px 0px 4px 4px',
            color: '#2B2B2B'
        });

        return (         
            <div className="container infoCardRoot" style={{
                    width: width,
                    height: height
                }}>
                <div className="row infoCardHeader" style={{
                    display:
                        (!icon || icon.trim() === '') &&
                        (!title || title.trim() === '') &&
                        (!href || href.trim() === '') ?
                        'none' : ''
                }}>
                    <div className="col" style={{ display:'flex', alignContent:'center', marginTop: '6px'}}>
                            <FontIcon
                                iconName={icon ? icon : ''}
                                className={iconClass}
                                hidden={!icon || icon === ''}
                            />
                        <Text variant='medium'
                            style={{ margin: '8px 0px 4px 4px', color: '#2B2B2B' }}
                                hidden={!title || title.toLowerCase().trim() === ''}
                            >
                                <b>&nbsp;{title}</b>
                            </Text>
                    </div>
                    <div className="col-3" style={{ textAlign: 'right', marginTop: '6px' }}>
                        <FontIcon
                            iconName='OpenInNewWindow'
                            className={popExpandIconClass}
                            hidden={!href || href.trim() === ''}
                            style={{ cursor: 'pointer' }}
                            title={hrefTitle ? hrefTitle : (href ? href : '')}
                            onClick={(event: React.MouseEvent<HTMLElement>) => { this._handleRedirect(href ? href : '') }}
                            />
                        </div>
                    </div>
                    <div className="row infoCardContent">
                        <div className="col">
                            {this.props.children}
                        </div>
                    </div>
                </div>
        );
    }

    componentDidUpdate(previousProps: IInfoCardState): void {
        if (this.props.title != previousProps.title ||
            this.props.height != previousProps.height ||
            this.props.width != previousProps.width)
        {
            this.setState({
                title: this.props.title,
                height: this.props.height,
                width: this.props.width
            });
        }
    }

    private _handleRedirect(url: string) {
        window.open(url, "_self");
    }
}