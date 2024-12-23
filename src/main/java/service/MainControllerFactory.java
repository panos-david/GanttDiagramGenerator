package service;

import app.ApplicationController;

public class MainControllerFactory {

    public IMainController createMainController() {
        return new ApplicationController();
    }
}
