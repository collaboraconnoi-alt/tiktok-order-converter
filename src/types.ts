import { Timestamp } from 'firebase/firestore';

export interface Message {
  id?: string;
  text: string;
  userId: string;
  userName: string;
  userPhoto?: string;
  timestamp: Timestamp;
  channelId: string;
}

export interface UserProfile {
  uid: string;
  displayName: string | null;
  photoURL: string | null;
  email: string | null;
}
