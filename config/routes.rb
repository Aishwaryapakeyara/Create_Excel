Rails.application.routes.draw do
  # get 'reports/index'
  # Define your application routes per the DSL in https://guides.rubyonrails.org/routing.html

  # Defines the root path route ("/")
  # root "articles#index"

  resources :reports, only: [:index]

  root 'reports#index'
end
