import * as React from 'react';
import { Text, IFontStyles } from '@fluentui/react';
import { TooltipHost } from '@fluentui/react';

interface IStatisticsTileState {
    value: number;
    title: string;
    oldValue?: number;
    fontSize?: keyof IFontStyles;
}

export class StatisticsTile extends React.Component<{ value: number, title: string, oldValue?: number, fontSize?: keyof IFontStyles }, IStatisticsTileState> {
    constructor(props: { value: number, title: string, oldValue?: number, fontSize?: keyof IFontStyles }) {
        super(props);

        this.state = {
            value: this.props.value,
            title: this.props.title,
            oldValue: this.props.oldValue,
            fontSize: this.props.fontSize
        };

    }

    public render() {
        const {
            value, title, oldValue, fontSize
        } = this.state;

        var increase = oldValue ? value - oldValue : 0;

        var percentage: number = oldValue ? (
            oldValue > 0 ? increase * 100 / oldValue : increase * 100
        ) : 0;

        percentage = Math.round(percentage * 100) / 100;

        var eventsIncreased: boolean = percentage > 0;
        percentage = percentage < 0 ? 0 - percentage : percentage;

        return (
            <div className="eventStatisticsTile" style={{ textAlign: 'center' }}>
                <Text variant={fontSize ? fontSize : 'superLarge'}>{value}</Text><br />
                <Text variant='smallPlus' style={{ position: 'relative', top: '-12px' }}>{title}</Text><br />
                <Text variant='xSmall' style={{
                    display: oldValue ? '' : 'none',
                    position: 'relative',
                    top: '-18px',
                    color: eventsIncreased ? 'red' : 'green'
                }}>
                    <TooltipHost
                        key={"eventStats_" + title}
                        content={ (increase === 0 ? "Same amount of events as" : 
                            (increase > 0 ? increase + (increase === 1 ? " event" : " events") + " more than"
                                : (0 - increase) + (increase === -1 ? " event" : " events") + " less than")) +
                            " in previous 30 days period"
                        }>
                        &nbsp;&nbsp;&nbsp;{percentage}% {increase === 0 ? "=" : (eventsIncreased ? "➚" : "➘")}
                    </TooltipHost>
                </Text>
            </div>

        );
    }

    componentDidMount() {
    }

    componentDidUpdate(prevProps: { value: number, title: string, oldValue?: number, fontSize?: keyof IFontStyles }) {
        if (prevProps.value != this.props.value ||
            prevProps.title != this.props.title ||
            prevProps.oldValue != this.props.oldValue ||
            prevProps.fontSize != this.props.fontSize) {

            this.setState({
                value: this.props.value,
                title: this.props.title,
                oldValue: this.props.oldValue,
                fontSize: this.props.fontSize
            });
        }
    }
}