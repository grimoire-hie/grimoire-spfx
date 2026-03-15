jest.mock('./GraphService', () => ({
  GraphService: jest.fn()
}));

import { GraphService } from './GraphService';
import { PeopleService, normalizePeopleSearchQuery } from './PeopleService';

describe('normalizePeopleSearchQuery', () => {
  it('removes conversational wrappers from people-card requests', () => {
    expect(normalizePeopleSearchQuery('show me Test User people card')).toBe('Test User');
    expect(normalizePeopleSearchQuery('who is Test User')).toBe('Test User');
  });
});

describe('PeopleService.searchPeople', () => {
  afterEach(() => {
    jest.restoreAllMocks();
  });

  it('prefers strong internal directory matches over loose people suggestions', async () => {
    const getSpy = jest.fn(async (path: string) => {
      const decodedPath = decodeURIComponent(path);

      if (decodedPath.startsWith('/users?')) {
        if (decodedPath.includes('startswith(displayName') || decodedPath.includes('startswith(givenName')) {
          return {
            success: true,
            data: {
              value: [
                {
                  id: 'user-1',
                  displayName: 'Test User',
                  givenName: 'Test',
                  surname: 'User',
                  mail: 'user@contoso.com',
                  userPrincipalName: 'user@contoso.com',
                  jobTitle: 'Solution Architect',
                  department: 'Digital Workplace',
                  officeLocation: 'Zurich',
                  businessPhones: ['+1 555 000 0000'],
                  userType: 'Member'
                }
              ]
            }
          };
        }

        return {
          success: true,
          data: { value: [] }
        };
      }

      if (decodedPath.startsWith('/me/people?')) {
        return {
          success: true,
          data: {
            value: [
              {
                id: 'people-1',
                displayName: 'User Nicole',
                scoredEmailAddresses: [
                  {
                    address: 'user@contoso.com',
                    relevanceScore: 0.91
                  }
                ]
              }
            ]
          }
        };
      }

      throw new Error(`Unexpected path: ${decodedPath}`);
    });
    (GraphService as jest.Mock).mockImplementation(() => ({
      get: getSpy
    }));

    const service = new PeopleService({} as never);
    const result = await service.searchPeople('Test User', 5);

    expect(getSpy).toHaveBeenCalled();
    expect(result.success).toBe(true);
    expect(result.data).toHaveLength(1);
    expect(result.data?.[0]).toEqual(expect.objectContaining({
      displayName: 'Test User',
      email: 'user@contoso.com',
      jobTitle: 'Solution Architect',
      department: 'Digital Workplace',
      officeLocation: 'Zurich',
      phone: '+1 555 000 0000'
    }));
  });
});
