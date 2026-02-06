export interface UserRole {
    [key: string]: any; // TODO: Refine this as we discover actual properties
}

export interface UserProfile {
    username: string;
    token: string;
    role: UserRole;
}
