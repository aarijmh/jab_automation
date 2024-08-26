class JabAction:
    action : str
    name : str
    role : str
    action_type : str
    value : any
    index_in_parent: int
    element_depth: int
    def __init__(self, action, name, role, action_type, value, index_in_parent, element_depth):
        self.action = action
        self.name = name
        self.role = role
        self.action_type = action_type
        self.value = value
        self.index_in_parent = index_in_parent
        self.element_depth = element_depth
        
    def __str__(self) -> str:
        return f"{self.name}  {self.role}  {self.action} {self.action_type} {self.value}"
    