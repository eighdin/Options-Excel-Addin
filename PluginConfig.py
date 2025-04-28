class PluginConfig:
    # MUTABLE CONFIG DEFINITIONS
    refresh_rate_mins: int # how many minutes between
    min_refresh_rate_mins: int # smallest value that refresh rate can be in minutes
    
    def __init__(self):
        # MUTABLE CONFIG VALUES
        self.refresh_rate_mins = 10
        self.min_refresh_rate_mins = 1 
        self.already_printed_logo = False
        
    # CONFIG FOR CONTRACT INFO OBJECT IS INSIDE THAT FILE, BECAUSE I'M LAZY