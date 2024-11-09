export interface AppSettings {
  clientId: string;
  clientSecret: string;
  tenantId: string;
}

const settings: AppSettings = {
  clientId: 'YOUR_CLIENT_ID_HERE',
  clientSecret: 'YOUR_CLIENT_SECRET_HERE',
  tenantId: 'YOUR_TENANT_ID_HERE',
};

export default settings;
