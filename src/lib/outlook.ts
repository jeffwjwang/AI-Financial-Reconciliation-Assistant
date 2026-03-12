import { db } from './db';

export interface OutlookTokens {
  access_token: string;
  refresh_token: string;
  expires_in: number;
}

export class OutlookService {
  private static async getAccessToken(): Promise<string | null> {
    const tokens = await db.settings.get('outlook_tokens');
    if (!tokens) return null;
    // In a real app, check expiration and use refresh_token
    return tokens.value.access_token;
  }

  static async fetchEmails(query: string, receivedAfter?: string) {
    const token = await this.getAccessToken();
    if (!token) throw new Error("Not authenticated with Outlook");

    let filter = `contains(subject, '${query}')`;
    if (receivedAfter) {
      filter += ` and receivedDateTime ge ${receivedAfter}`;
    }

    const url = `https://graph.microsoft.com/v1.0/me/messages?$filter=${encodeURIComponent(filter)}&$expand=attachments&$top=50`;
    console.log('Fetching emails with URL:', url);
    const response = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` }
    });

    if (!response.ok) throw new Error("Failed to fetch emails");
    const data = await response.json();
    return data.value || [];
  }

  static async getAttachment(messageId: string, attachmentId: string) {
    const token = await this.getAccessToken();
    if (!token) throw new Error("Not authenticated");

    const response = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${messageId}/attachments/${attachmentId}`, {
      headers: { Authorization: `Bearer ${token}` }
    });

    if (!response.ok) throw new Error("Failed to fetch attachment");
    const data = await response.json();
    return data.contentBytes; // Base64
  }
}
