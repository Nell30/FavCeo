import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import styles from './FavoriteCeoQuestionWebPart.module.scss';
import { SPHttpClient } from '@microsoft/sp-http';


export interface IFavoriteCeoQuestionWebPartProps {
  description: string;
}
export interface ISPList {
  Id: number;
  Created: string;
  Answers: string;
  Better: string;
  Replies: string;
  Status: string;
}

export default class FavoriteCeoQuestionWebPart extends BaseClientSideWebPart<IFavoriteCeoQuestionWebPartProps> {

  public async render(): Promise<void> {
    const favorites = await this.getTopFavorites();
    try{
      this.domElement.innerHTML = `

      ${favorites.length > 0 ? this.renderSlideshow(favorites) : '<p>No favorite questions found.</p>'}
        
      <p>Interested in asking the CEO a question? <a href="https://torgrace.sharepoint.com/:u:/r/sites/Nelson/SitePages/Share-Your-Thoughts-With-The-CEO.aspx?csf=1&web=1&share=EScBSWpHQGBJpNf9CbpIsUkBDxSwVrMY-UCd43rbm4Bb4g&e=lCrZHc" target="_blank">Click here to share your thoughts!</a></p>
      `;

      if (favorites.length > 0) {
        this.setupSlideshow();
      }

      

    }catch (error) {
      console.error('Error rendering AskCeoWebPart:', error);
    }
  }

  private setupSlideshow(): void {
    const slides = this.domElement.querySelectorAll(`.${styles.slide}`) as NodeListOf<HTMLElement>;
    const prevButton = this.domElement.querySelector(`.${styles.slidePrev}`) as HTMLElement;
    const nextButton = this.domElement.querySelector(`.${styles.slideNext}`) as HTMLElement;
    const slideshowContainer = this.domElement.querySelector(`.${styles.slideshowContainer}`) as HTMLElement;
  
    if (slides.length === 0) {
      if (slideshowContainer) {
        slideshowContainer.innerHTML = '<p>No favorite questions found.</p>';
      }
      if (prevButton) prevButton.style.display = 'none';
      if (nextButton) nextButton.style.display = 'none';
      return;
    }
  
    let currentSlide = 0;
  
    const adjustContainerHeight = (slideIndex: number) => {
      const slideHeight = slides[slideIndex].scrollHeight;
      slideshowContainer.style.height = `${slideHeight}px`;
    };
  
    const showSlide = (n: number) => {
      slides[currentSlide].classList.remove(styles.active);
      currentSlide = (n + slides.length) % slides.length;
      slides[currentSlide].classList.add(styles.active);
      adjustContainerHeight(currentSlide);
    };
  
    prevButton?.addEventListener('click', () => showSlide(currentSlide - 1));
    nextButton?.addEventListener('click', () => showSlide(currentSlide + 1));
  
    // Initial height adjustment
    adjustContainerHeight(currentSlide);
  
    // Auto-advance slides every 10 seconds
    setInterval(() => showSlide(currentSlide + 1), 10000);
  
    // Adjust height on window resize
    window.addEventListener('resize', () => adjustContainerHeight(currentSlide));
  }

  private renderSlideshow(favorites: ISPList[]): string {
    let slideshowHtml = `
      <div class="${styles.slideshow}">
        <div class="${styles.slideshowTitle}">Top Questions From The CEO</div>
        <div class="${styles.slideshowContainer}">
    `;

    favorites.forEach((item, index) => {
      slideshowHtml += `
        <div class="${styles.slide} ${index === 0 ? styles.active : ''}" style="height: auto; overflow: auto;">
          <div class="${styles.slideContent}">
            <div class="${styles.card}">
              <h3 style="color:#005596;">Q: ${item.Answers}</h3>
              <hr width="100%">
                <h4 class="${styles.answerHeader}">Answer:</h4>
                <p class="${styles.answerText}">${item.Replies}</p>
              <p class="${styles.submissionDate}">Submitted: ${new Date(item.Created).toLocaleDateString()}</p>
            </div>
          </div>
        </div>
      `;
    });

    slideshowHtml += `
        </div>
        <button class="${styles.slidePrev}">❮</button>
        <button class="${styles.slideNext}">❯</button>
      </div>
    `;

    return slideshowHtml;
  }


  private async getTopFavorites(): Promise<ISPList[]> {
    try {
      console.log('Fetching top favorites...');
      
      // Check for IsFavorite Yes (1)
      let url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Reflection')/items?$filter=IsFavorite eq 1&$top=10`;
      console.log('IsFavorite Yes URL:', url);
      let response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      let data = await response.json();
      console.log('IsFavorite Yes results:', data.value);
  
      // Check for IsFavorite No (0)
      url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Reflection')/items?$filter=IsFavorite eq 0&$top=10`;
      console.log('IsFavorite No URL:', url);
      response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      data = await response.json();
      console.log('IsFavorite No results:', data.value);
  
      // Check for all items and their IsFavorite values
      url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Reflection')/items?$select=Id,Title,IsFavorite,Status&$top=10`;
      console.log('All items URL:', url);
      response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      data = await response.json();
      console.log('All items results:', data.value);
  
      // Updated original query
      url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Reflection')/items?$filter=Status eq 'Approved' and IsFavorite eq 1&$orderby=Created desc&$top=10`;
      console.log('Updated original query URL:', url);
      response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      data = await response.json();
      console.log('Updated original query results:', data.value);
  
      return data.value;
    } catch (error) {
      console.error('Error fetching top favorites:', error);
      return [];
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

}
