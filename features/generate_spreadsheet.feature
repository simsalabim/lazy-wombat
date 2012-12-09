Feature: As a lazy wombat I should be able to generate an xlsx spreadsheet

  Scenario: I generate spreadsheet and save it to file
    Given I generate following spreadsheet:
      | Arnold  | Schwarzenegger | Cyberdyne Systems Model 101 | The Terminator    | 1       |
      | Wombat  | Animal         | Universal                   | Optimum Nutrition | 001     |
      | Hamster | Mole           | Apartments                  | Valuation         | Mockups |
    And save it as "my_spreadsheet.xlsx"
    Then "my_spreadsheet.xlsx" file should exist
    And "my_spreadsheet.xlsx" should have 1 spreadsheet
    And "my_spreadsheet.xlsx" should look like this:
      | Arnold  | Schwarzenegger | Cyberdyne Systems Model 101 | The Terminator    | <1.0>   |
      | Wombat  | Animal         | Universal                   | Optimum Nutrition | <1.0>   |
      | Hamster | Mole           | Apartments                  | Valuation         | Mockups |
