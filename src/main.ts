import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';
import { ElementModule } from './app/element.module';

platformBrowserDynamic().bootstrapModule(ElementModule)
    .catch(err => console.error(err));
