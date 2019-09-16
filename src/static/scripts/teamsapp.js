(function() {
    'use strict';
    // Call the initialize API first
    microsoftTeams.initialize();
   
    // Check the initial theme user chose and respect it
    microsoftTeams.getContext(function(context){
        if (context && context.theme) {
            setTheme(context.theme);
        }
    });

    // Handle theme changes
    microsoftTeams.registerOnThemeChangeHandler(function(theme) {
        setTheme(theme);
    });

    // Save configuration changes
    microsoftTeams.settings.registerOnSaveHandler(function(saveEvent) {
        // Let the Microsoft Teams platform know what you want to load based on
        // what the user configured on this page
        microsoftTeams.settings.setSettings({
            contentUrl: createTabUrl(), // Mandatory parameter
            entityId: createTabUrl() // Mandatory parameter
        });

        // Tells Microsoft Teams platform that we are done saving our settings. Microsoft Teams waits
        // for the app to call this API before it dismisses the dialog. If the wait times out, you will
        // see an error indicating that the configuration settings could not be saved.
        console.log(saveEvent.notifySuccess());
        saveEvent.notifySuccess();
    });

    // Logic to let the user configure what they want to see in the tab being loaded
    document.addEventListener('DOMContentLoaded', function() {
        console.log('something')
        var tabChoice = document.getElementById('tabChoice');
        console.log('tabchoice: ',tabChoice)
        if (tabChoice) {
            console.log('tabchoise is true')
            tabChoice.onchange = function() {
                var selectedTab = this[this.selectedIndex].value;

                // This API tells Microsoft Teams to enable the 'Save' button. Since Microsoft Teams always assumes
                // an initial invalid state, without this call the 'Save' button will never be enabled.
                microsoftTeams.settings.setValidityState(selectedTab === 'first' || selectedTab === 'second');
            };
        }
    });

    // Set the desired theme
    function setTheme(theme) {
        if (theme) {
            // Possible values for theme: 'default', 'light', 'dark' and 'contrast'
            document.body.className = 'theme-' + (theme === 'default' ? 'light' : theme);
        }
    }

    // Create the URL that Microsoft Teams will load in the tab. You can compose any URL even with query strings.
    function createTabUrl() {
        console.log('inside createTabUrl')
        var tabChoice = document.getElementById('tabChoice');
        consol.log('tabchoice: ',tabChoice)
        var selectedTab = tabChoice[tabChoice.selectedIndex].value;
        consol.log('selectedTab: ',selectedTab)
        consol.log('window.location.protocol: ',window.location.protocol)
        consol.log('window.location.host: ',window.location.host)

        return window.location.protocol + '//' + window.location.host + '/' + selectedTab;
    }
})();
