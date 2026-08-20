import type {ReactNode} from 'react';
import clsx from 'clsx';
import Heading from '@theme/Heading';
import styles from './styles.module.css';

type FeatureItem = {
  title: string;
  Svg: React.ComponentType<React.ComponentProps<'svg'>>;
  description: ReactNode;
};

const FeatureList: FeatureItem[] = [
  {
    title: 'Works in the web browser and Excel',
    Svg: require('@site/static/img/undraw_docusaurus_mountain.svg').default,
    description: (
      <>
        The Excel Control Charts add-in is a web application that can be used in the web browser or in Microsoft Excel. 
        It allows users to create control charts and funnel charts from their data without needing to install any additional software.
      </>
    ),
  },
  {
    title: 'Open source',
    Svg: require('@site/static/img/undraw_docusaurus_tree.svg').default,
    description: (
      <>
        The code is available on <a href="https://github.com/AUS-DOH-Safety-and-Quality/ExcelControlCharts">GitHub</a> and free to use under the GPL-3.0 license.
      </>
    ),
  },
  {
    title: 'Built upon PowerBI-SPC and PowerBI-Funnels',
    Svg: require('@site/static/img/undraw_docusaurus_react.svg').default,
    description: (
      <>
        Extends the functionality of PowerBI-SPC (<a href="https://github.com/AUS-DOH-Safety-and-Quality/PowerBI-SPC">GitHub</a>) and PowerBI-Funnels (<a href="https://github.com/AUS-DOH-Safety-and-Quality/PowerBI-Funnels">GitHub</a>) for users that want to create these charts from the web browser or Excel.
      </>
    ),
  },
];

function Feature({title, Svg, description}: FeatureItem) {
  return (
    <div className={clsx('col col--4')}>
      <div className="text--center">
        <Svg className={styles.featureSvg} role="img" />
      </div>
      <div className="text--center padding-horiz--md">
        <Heading as="h3">{title}</Heading>
        <p>{description}</p>
      </div>
    </div>
  );
}

export default function HomepageFeatures(): ReactNode {
  return (
    <section className={styles.features}>
      <div className="container">
        <div className="row">
          {FeatureList.map((props, idx) => (
            <Feature key={idx} {...props} />
          ))}
        </div>
      </div>
    </section>
  );
}
