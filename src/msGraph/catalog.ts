import { graphFetch } from './auth.js';
import { logger } from '../logger.js';

type CatalogLinks = {
  repos: string[];
  base44: string[];
  dataBuckets: string[];
};

function buildHtmlSection(title: string, links: string[]): string {
  const items = links.length
    ? links.map((link) => `<li><a href="${link}">${link}</a></li>`).join('')
    : '<li><em>No links provided</em></li>';
  return `<div data-title="${title}"><h2>${title}</h2><ul>${items}</ul></div>`;
}

export async function ensureCatalogPage(siteId: string, title: string, links: CatalogLinks): Promise<string> {
  const html = [
    buildHtmlSection('Repos', links.repos),
    buildHtmlSection('Base44 Apps', links.base44),
    buildHtmlSection('Data Buckets', links.dataBuckets)
  ].join('');

  const response = await graphFetch(`/sites/${siteId}/pages`, {
    method: 'POST',
    body: JSON.stringify({
      '@odata.type': '#microsoft.graph.sitePage',
      name: 'Catalog.aspx',
      title,
      publishingState: {
        level: 'published',
        versionId: '0.1'
      },
      canvasLayout: {
        horizontalSections: [
          {
            layout: 'threeColumns',
            columns: [
              { width: 4, emphasis: 'none', webparts: [] },
              { width: 4, emphasis: 'none', webparts: [] },
              { width: 4, emphasis: 'none', webparts: [] }
            ]
          }
        ]
      },
      webParts: [
        {
          type: 'text',
          data: {
            innerHtml: html
          }
        }
      ]
    })
  });

  logger.info('Catalog page ensured', { pageUrl: response?.webUrl });
  return response?.webUrl as string;
}
