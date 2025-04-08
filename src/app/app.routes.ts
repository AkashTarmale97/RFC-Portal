import { Routes } from '@angular/router';
import { ChangeRequestComponent } from './change-request/change-request.component';
import { MainPageComponent } from './main-page/main-page.component';

export const routes: Routes = [
    {
        path: "",
        pathMatch:'full',
        redirectTo:"index"

    },
    {
        path: "change-request",
        component:ChangeRequestComponent
    },
    {
        path: "index",
        component:MainPageComponent
    }
];
