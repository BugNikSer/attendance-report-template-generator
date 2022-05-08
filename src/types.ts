export type IGroup = Array<string>;

export interface ITeacher {
    [group: string]: IGroup;
}

export interface IStudents {
    [teacher: string]: ITeacher;
}
