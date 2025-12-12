import axios from 'axios';

interface EmailMessage {
  id: string;
  subject: string;
  from: {
    emailAddress: {
      name: string;
      address: string;
    };
  };
  receivedDateTime: string;
  bodyPreview: string;
  body?: {
    contentType: string;
    content: string;
  };
  hasAttachments: boolean;
  isRead: boolean;
  importance: string;
}

interface EmailListResponse {
  value: EmailMessage[];
  '@odata.nextLink'?: string;
}

export class OutlookClient {
  private baseUrl = 'https://graph.microsoft.com/v1.0';
  private accessToken: string | null = null;

  setAccessToken(token: string) {
    this.accessToken = token;
  }

  private getHeaders() {
    if (!this.accessToken) {
      throw new Error('Access token is required. Please obtain an access token from Microsoft Graph Explorer: https://developer.microsoft.com/en-us/graph/graph-explorer');
    }
    return {
      'Authorization': `Bearer ${this.accessToken}`,
      'Content-Type': 'application/json'
    };
  }

  async listEmails(params: {
    top?: number;
    skip?: number;
    filter?: string;
    orderBy?: string;
    select?: string;
    search?: string;
  } = {}): Promise<EmailMessage[]> {
    try {
      const queryParams = new URLSearchParams();

      if (params.top) queryParams.append('$top', params.top.toString());
      if (params.skip) queryParams.append('$skip', params.skip.toString());
      if (params.filter) queryParams.append('$filter', params.filter);
      if (params.orderBy) queryParams.append('$orderby', params.orderBy);
      if (params.select) queryParams.append('$select', params.select);
      if (params.search) queryParams.append('$search', `"${params.search}"`);

      const url = `${this.baseUrl}/me/messages?${queryParams.toString()}`;
      const response = await axios.get<EmailListResponse>(url, {
        headers: this.getHeaders()
      });

      return response.data.value;
    } catch (error: any) {
      if (error.response?.status === 401) {
        throw new Error('Invalid or expired access token. Please obtain a new token from Microsoft Graph Explorer: https://developer.microsoft.com/en-us/graph/graph-explorer');
      }
      throw new Error(`Failed to list emails: ${error.message}`);
    }
  }

  async getEmail(messageId: string): Promise<EmailMessage> {
    try {
      const url = `${this.baseUrl}/me/messages/${messageId}`;
      const response = await axios.get<EmailMessage>(url, {
        headers: this.getHeaders()
      });

      return response.data;
    } catch (error: any) {
      if (error.response?.status === 401) {
        throw new Error('Invalid or expired access token. Please obtain a new token from Microsoft Graph Explorer: https://developer.microsoft.com/en-us/graph/graph-explorer');
      }
      if (error.response?.status === 404) {
        throw new Error(`Email with ID ${messageId} not found.`);
      }
      throw new Error(`Failed to get email: ${error.message}`);
    }
  }

  async markAsRead(messageId: string, isRead: boolean = true): Promise<void> {
    try {
      const url = `${this.baseUrl}/me/messages/${messageId}`;
      await axios.patch(url, { isRead }, {
        headers: this.getHeaders()
      });
    } catch (error: any) {
      if (error.response?.status === 401) {
        throw new Error('Invalid or expired access token. Please obtain a new token from Microsoft Graph Explorer: https://developer.microsoft.com/en-us/graph/graph-explorer');
      }
      throw new Error(`Failed to mark email as read: ${error.message}`);
    }
  }

  async deleteEmail(messageId: string): Promise<void> {
    try {
      const url = `${this.baseUrl}/me/messages/${messageId}`;
      await axios.delete(url, {
        headers: this.getHeaders()
      });
    } catch (error: any) {
      if (error.response?.status === 401) {
        throw new Error('Invalid or expired access token. Please obtain a new token from Microsoft Graph Explorer: https://developer.microsoft.com/en-us/graph/graph-explorer');
      }
      throw new Error(`Failed to delete email: ${error.message}`);
    }
  }

  async searchEmails(searchQuery: string, top: number = 10): Promise<EmailMessage[]> {
    return this.listEmails({
      search: searchQuery,
      top,
      orderBy: 'receivedDateTime DESC'
    });
  }
}