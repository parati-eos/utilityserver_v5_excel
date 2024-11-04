export interface IScraper {
  scrapeParagraphs(url: string): Promise<{
    paragraphs: string;
    executionTime: number;
  }>;
}
