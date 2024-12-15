import * as React from 'react';
import { Stack, Text, PrimaryButton, DefaultButton } from '@fluentui/react';

interface ImmoMailScreenProps {
  login: () => void;
}

const ImmoMailScreen: React.FC<ImmoMailScreenProps> = ({ login }) => {
  return (
    <Stack
      verticalAlign="center"
      horizontalAlign="center"
      styles={{
        root: {
          height: '100vh',
          backgroundColor: '#ffffff',
          position: 'relative', // allows us to position the feedback button at the bottom right
        },
      }}
      tokens={{ childrenGap: 20 }}
    >
      {/* Icon, Title, and Paragraph in a single vertical stack */}
      <Stack verticalAlign="center" horizontalAlign="center" tokens={{ childrenGap: 15 }}>
        {/* Icon */}
        <svg
          width="70"
          height="70"
          viewBox="0 0 132 125"
          fill="none"
          xmlns="http://www.w3.org/2000/svg"
        >
          <g clipPath="url(#clip0_99_80)">
            <path
              d="M124.592 122.315H7.40816C4.80857 122.315 2.69388 120.207 2.69388 117.616V50.7599C2.69388 49.3234 3.35388 47.954 4.49878 47.068L57.1641 5.73247C62.3498 1.6647 69.6637 1.6647 74.8494 5.73247L127.515 47.0546C128.646 47.9406 129.32 49.31 129.32 50.7464V117.603C129.32 120.194 127.205 122.302 124.605 122.302L124.592 122.315Z"
              fill="white"
              stroke="#231F20"
              strokeWidth="4"
              strokeMiterlimit="10"
            />
            <path
              d="M4.04082 49.82L52.5306 86.0675C54.7531 87.1146 58.4976 88.5243 63.3061 88.7525C69.5559 89.0478 74.378 87.1952 76.7755 86.0675C87.5376 79.0328 107.755 63.1108 126.612 48.4775"
              fill="white"
            />
            <path
              d="M4.04082 49.82L52.5306 86.0675C54.7531 87.1146 58.4976 88.5243 63.3061 88.7525C69.5559 89.0478 74.378 87.1952 76.7755 86.0675C87.5376 79.0328 107.755 63.1108 126.612 48.4775"
              stroke="#231F20"
              strokeWidth="4"
              strokeMiterlimit="10"
            />
            <path
              d="M5.38776 119.63L57.9184 87.4102"
              stroke="#231F20"
              strokeWidth="4"
              strokeMiterlimit="10"
            />
            <path
              d="M74.0816 87.4102L127.959 120.973"
              stroke="#231F20"
              strokeWidth="4"
              strokeMiterlimit="10"
            />
            <path
              d="M22.898 8.20264V31.0251"
              stroke="#231F20"
              strokeWidth="4"
              strokeMiterlimit="10"
            />
            <path
              d="M39.0612 6.86011V20.2851"
              stroke="#231F20"
              strokeWidth="4"
              strokeMiterlimit="10"
            />
            <path
              d="M20.2041 6.86011H39.0408"
              fill="white"
              stroke="#231F20"
              strokeWidth="4"
              strokeMiterlimit="10"
            />
            <path
              d="M4.04082 49.82L126.612 48.4775"
              stroke="#231F20"
              strokeWidth="4"
              strokeMiterlimit="10"
            />
          </g>
          <defs>
            <clipPath id="clip0_99_80">
              <rect width="132" height="125" fill="white" />
            </clipPath>
          </defs>
        </svg>

        {/* Title */}
        <Text variant="xLarge" styles={{ root: { fontWeight: 'bold' } }}>
          Willkommen bei <span style={{ color: '#231F20' }}>ImmoMail</span>
        </Text>

        {/* Paragraph Text */}
        <Text variant="medium" styles={{ root: { textAlign: 'center', fontSize: '14px' } }}>
          Wir helfen dir Mieteranfragen zu beantworten. Schnell und unkompliziert
        </Text>
      </Stack>

      {/* Connect Button */}
      <PrimaryButton
        text="Mit Office 360 Verbinden"
        onClick={login}
        styles={{
          root: {
            backgroundColor: '#000000',
            borderColor: '#000000',
            color: '#ffffff',
            fontSize: '14px',
            padding: '10px 20px',
            borderRadius: '5px',
            width: '240px',
          },
          rootHovered: {
            backgroundColor: '#1a1a1a',
            borderColor: '#1a1a1a',
          },
        }}
      />

      {/* Feedback Button at the bottom right */}
      <DefaultButton
        text="Feedback"
        onClick={() => {
          // handle feedback button click
          alert('Feedback button clicked!');
        }}
        styles={{
          root: {
            position: 'absolute',
            bottom: '20px',
            right: '20px',
            backgroundColor: '#f3f2f1',
            borderColor: '#f3f2f1',
            color: '#000000',
            fontSize: '14px',
            padding: '5px 15px',
            borderRadius: '4px',
          },
          rootHovered: {
            backgroundColor: '#e1dfdd',
            borderColor: '#e1dfdd',
          },
        }}
      />
    </Stack>
  );
};

export default ImmoMailScreen;
