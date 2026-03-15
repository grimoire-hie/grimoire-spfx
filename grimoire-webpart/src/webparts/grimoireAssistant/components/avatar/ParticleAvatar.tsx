/**
 * ParticleAvatar
 * Renderer adapter.
 * Uses pure SVG renderer.
 */

import * as React from 'react';
import { ISvgAvatarProps, SvgAvatar } from './SvgAvatar';

export interface IParticleAvatarProps extends ISvgAvatarProps {}

const ParticleAvatarInner: React.FC<IParticleAvatarProps> = (props) => <SvgAvatar {...props} />;

ParticleAvatarInner.displayName = 'ParticleAvatar';

export const ParticleAvatar = React.memo(ParticleAvatarInner);
